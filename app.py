import io
import re
import csv
import docx
import pandas as pd
import streamlit as st

def normalize_text(text):
    """Cleans up spaces and non-standard line breaks from Word formatting."""
    if not text:
        return ""
    # Replace non-breaking spaces
    clean_text = text.replace('\xa0', ' ').replace('\u202f', ' ')
    # Normalize multiple line breaks
    clean_text = re.sub(r'[\r\x0b\x0c\n]+', '\n', clean_text)
    # Normalize multiple spaces
    clean_text = re.sub(r'[ \t]+', ' ', clean_text)
    return clean_text.strip()

def extract_value(line):
    """Extracts the string portion located after the first ':' character."""
    if ":" in line:
        return line.split(":", 1)[1].strip()
    return ""

def clean_for_csv(text, separator=" "):
    """Replaces line breaks with a separator to avoid breaking CSV formatting."""
    if not text:
        return ""
    # Replace carriage returns and newlines
    text_no_nl = text.replace('\n', separator).replace('\r', separator)
    # Remove duplicate spaces
    return re.sub(r'\s+', ' ', text_no_nl).strip()

def extract_lines_from_document(doc_file):
    """Extracts lines while preventing duplication from merged cells safely."""
    doc = docx.Document(doc_file)
    global_lines = []
    skipped_logs = []
    
    for idx, table in enumerate(doc.tables):
        # Track unique cell elements to prevent repeating text from merged cells
        processed_cells = set()
        table_lines = []
        
        for r_idx, row in enumerate(table.rows):
            for c_idx, cell in enumerate(row.cells):
                # Check underlying XML element to identify merged duplicates
                if cell._tc in processed_cells:
                    continue
                processed_cells.add(cell._tc)
                
                cell_text = normalize_text(cell.text)
                if cell_text:
                    # Deduplicate repeating lines inside the SAME cell only
                    cell_lines = []
                    for line in cell_text.split('\n'):
                        line_strip = line.strip()
                        if line_strip and line_strip not in cell_lines:
                            cell_lines.append(line_strip)
                            table_lines.append((line_strip, f"Table {idx+1}, Row {r_idx+1}, Cell {c_idx+1}"))
        
        # Verify if the table contains inventory cards
        complete_table_text = " ".join([l[0] for l in table_lines]).lower()
        if "numéro d'inventaire" not in complete_table_text and "numero d'inventaire" not in complete_table_text:
            preview = " | ".join([l[0] for l in table_lines[:3]])
            skipped_logs.append(f"Table #{idx+1} skipped (Preview: {preview[:100]}...)")
        
        global_lines.extend(table_lines)
        
    return global_lines, skipped_logs

def parse_line_stream(lines):
    """Parses extracted lines and separates objects based on the inventory number key."""
    objects = []
    current_obj = None
    current_section = None
    
    for line_tuple in lines:
        line, source = line_tuple
        line_lower = line.lower()
        
        is_new_trigger = False
        if "numéro d'inventaire" in line_lower or "numero d'inventaire" in line_lower:
            if ":" in line:
                value_key = extract_value(line)
                if value_key:
                    is_new_trigger = True
                    
        if is_new_trigger:
            if current_obj is not None:
                objects.append(current_obj)
                
            current_obj = {
                "Inventory": extract_value(line),
                "Designations": [],
                "Function_Role": "",
                "Materials": [],
                "Height": "", 
                "Width": "", 
                "Thickness": "", 
                "Description": "",
                "Domain": [],
                "Acquisition": "",
                "Bibliography": "",
                "Notes": ""
            }
            current_section = None
            continue
            
        if current_obj is None:
            continue
            
        # Parse fields
        if "désignation du bien" in line_lower or "designation du bien" in line_lower:
            val = extract_value(line)
            if val and val not in current_obj["Designations"]:
                current_obj["Designations"].append(val)
            current_section = None
            
        elif "fonction / rôle" in line_lower or "fonction / role" in line_lower or "fonction/rôle" in line_lower:
            val = extract_value(line)
            if val:
                current_obj["Function_Role"] = val
            current_section = "Function_Role"
            
        elif "matière :" in line_lower or "matiere :" in line_lower:
            val = extract_value(line)
            if val:
                current_obj["Materials"].append(val.capitalize())
            current_section = "Materials"
            
        elif "hauteur en cm" in line_lower:
            current_obj["Height"] = extract_value(line)
            current_section = None
        elif "largeur en cm" in line_lower:
            current_obj["Width"] = extract_value(line)
            current_section = None
        elif "epaisseur en cm" in line_lower or "épaisseur en cm" in line_lower or "profondeur en cm" in line_lower or "diametre en cm" in line_lower or "diamètre en cm" in line_lower:
            current_obj["Thickness"] = extract_value(line)
            current_section = None
            
        elif "description analytique" in line_lower:
            current_obj["Description"] = extract_value(line)
            current_section = "Description"
            
        elif "domaine :" in line_lower:
            val = extract_value(line)
            if val:
                current_obj["Domain"].append(val)
            current_section = "Domain"
            
        elif "acquisition" in line_lower:
            current_obj["Acquisition"] = extract_value(line)
            current_section = None
        elif "référence bibliographique" in line_lower or "reference bibliographique" in line_lower:
            current_obj["Bibliography"] = extract_value(line)
            current_section = None
        elif "notes :" in line_lower:
            current_obj["Notes"] = extract_value(line)
            current_section = None
            
        # Multi-line text accumulation logic
        else:
            if ":" not in line:
                if current_section == "Materials":
                    current_obj["Materials"].append(line.capitalize())
                elif current_section == "Description":
                    current_obj["Description"] += " " + line
                elif current_section == "Domain":
                    current_obj["Domain"].append(line)
                elif current_section == "Function_Role":
                    # For multi-line values under Function/Role
                    if current_obj["Function_Role"]:
                        current_obj["Function_Role"] += ", " + line
                    else:
                        current_obj["Function_Role"] = line
            else:
                current_section = None
                
    if current_obj is not None:
        objects.append(current_obj)
        
    return objects

def convert_to_dataframe(objects):
    """Converts parsed objects list to a structured pandas DataFrame."""
    headers = [
        "Inventaire", "Désignation du bien 1", "Désignation du bien 2", 
        "Fonction/rôle", "Fabricant", "Lieu de fabrication", "Date de fabrication", 
        "Date d'utilisation", "Lieu d'utilisation", "Matieres", "Longueur", 
        "Largeur", "Hauteur", "Description", "Domaine", "Mode d'acquisition", 
        "Date collecte", "Lieu de collecte", "Donateur", "Collecteur", 
        "Inscription", "Description restauration", "Date de restauration", 
        "Références bibliographiques", "Notes"
    ]
    
    rows = []
    for obj in objects:
        desig1 = clean_for_csv(obj["Designations"][0]) if len(obj["Designations"]) > 0 else ""
        desig2 = clean_for_csv(obj["Designations"][1]) if len(obj["Designations"]) > 1 else ""
        
        materials_joined = clean_for_csv(", ".join(obj["Materials"]))
        domain_joined = clean_for_csv(", ".join(obj["Domain"]))
        description_clean = clean_for_csv(obj["Description"])
        
        biblio_complete = obj["Bibliography"]
        if obj["Notes"]:
            biblio_complete += f" p. {obj['Notes']}" if biblio_complete else obj["Notes"]
        biblio_complete = clean_for_csv(biblio_complete)
        
        row = [
            clean_for_csv(obj["Inventory"]),
            desig1,
            desig2,
            clean_for_csv(obj["Function_Role"]),
            "", "", "", "", "", # Placeholder empty columns
            materials_joined,
            clean_for_csv(obj["Height"]),    # Hauteur en cm -> Longueur column
            clean_for_csv(obj["Width"]),     # Largeur en cm -> Largeur column
            clean_for_csv(obj["Thickness"]), # Epaisseur/Profondeur/Diamètre en cm -> Hauteur column
            description_clean,
            domain_joined,
            clean_for_csv(obj["Acquisition"]),
            "", "", "", "", "", "", "", # Placeholder empty columns
            biblio_complete,
            ""
        ]
        rows.append(row)
        
    return pd.DataFrame(rows, columns=headers)

# Streamlit UI
st.set_page_config(page_title="Document Inventory Parser", layout="wide")

st.title("Document Inventory Parser")
st.write("Upload a Word document (`.docx`) containing inventory tables to convert it into a structured CSV file.")

uploaded_file = st.file_uploader("Choose a Word document", type="docx")

if uploaded_file is not None:
    st.info("Processing file, please wait...")
    
    # Run processing
    lines, logs = extract_lines_from_document(uploaded_file)
    parsed_objects = parse_line_stream(lines)
    
    if parsed_objects:
        df = convert_to_dataframe(parsed_objects)
        
        # Display summary statistics
        st.success("Extraction complete!")
        col1, col2 = st.columns(2)
        col1.metric("Objects Extracted", len(parsed_objects))
        col2.metric("Tables Skipped", len(logs))
        
        # Warning logs for skipped tables
        if logs:
            with st.expander("View Processing Details"):
                for log in logs:
                    st.text(log)
                    
        # Preview extracted data
        st.subheader("Data Preview")
        st.dataframe(df.head(10))
        
        # Generate downloadable CSV content
        csv_buffer = io.StringIO()
        df.to_csv(csv_buffer, sep=';', index=False, quoting=csv.QUOTE_MINIMAL)
        csv_data = csv_buffer.getvalue()
        
        st.download_button(
            label="Download CSV file",
            data=csv_data,
            file_name="extracted_inventory.csv",
            mime="text/csv"
        )
    else:
        st.error("No valid inventory objects detected. Please verify that the tables in the document contain the expected fields.")