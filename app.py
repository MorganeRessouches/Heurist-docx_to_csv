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
    current_block = "general"
    
    # Major section headers that break multi-line note and text accumulation
    major_headers = [
        "identification :", "désignation :", "création/exécution :", 
        "epoque, datation :", "époque, datation :", "lieu :",
        "matière et technique :", "mesures :", "inscriptions / marques :", 
        "description analytique", "domaine :", "statut administratif :", 
        "acquisition", "bibliographie :", "multimédia :", "fonction d'usage :",
        "type d'util. / dest. :", "utilisation / destination :", "date d'util. / dest. :", "lieu d'util. / dest. :",
        "collecte :", "lieu de collecte :", "collecteur :", "date de collecte :", "méthode de collecte :",
        "type d'inscription :", "emplacement :", "transcription :"
    ]
    
    for line_tuple in lines:
        line, source = line_tuple
        line_lower = line.lower()
        
        # 1. Block Context Tracking
        if "bibliographie :" in line_lower or "reference bibliographique" in line_lower:
            current_block = "bibliography"
        elif "création/exécution :" in line_lower or "creation/execution :" in line_lower:
            current_block = "creation_execution"
        elif "fonction d'usage :" in line_lower:
            current_block = "utilisation"
        elif "collecte :" in line_lower:
            current_block = "collecte"
        elif "inscriptions / marques :" in line_lower or "inscriptions/marques :" in line_lower:
            current_block = "inscriptions"
        elif any(k in line_lower for k in ["identification :", "désignation :", "mesures :", "description analytique"]):
            current_block = "general"
            
        # 2. Exit General Notes mode if we hit a major section header
        if any(h in line_lower for h in major_headers):
            if current_section == "General_Notes":
                current_section = None
                
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
                "Biblio_Notes": "", 
                "General_Notes": [],
                "Date_Fabrication": "",
                "Lieu_Fabrication": "",
                "Date_Utilisation": "",
                "Lieu_Utilisation": "",
                "Lieu_Collecte": "",
                "Collecteur": "",
                "Date_Collecte": "",
                "Inscriptions": []
            }
            current_section = None
            current_block = "general"
            continue
            
        if current_obj is None:
            continue
            
        # 3. If locked inside a general note, capture raw text safely
        if current_section == "General_Notes":
            current_obj["General_Notes"].append(line)
            continue
            
        # 4. Standard Field Parsers
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
            
        # Fabrication Date & Place Rules
        elif "epoque, datation" in line_lower or "époque, datation" in line_lower:
            current_obj["Date_Fabrication"] = extract_value(line)
            current_section = "Date_Fabrication"
        elif "lieu :" in line_lower and current_block == "creation_execution":
            current_obj["Lieu_Fabrication"] = extract_value(line)
            current_section = "Lieu_Fabrication"
            
        # Utilisation Rules
        elif "date d'util. / dest." in line_lower:
            current_obj["Date_Utilisation"] = extract_value(line)
            current_section = "Date_Utilisation"
        elif "lieu d'util. / dest." in line_lower:
            current_obj["Lieu_Utilisation"] = extract_value(line)
            current_section = "Lieu_Utilisation"
            
        # Collecte Rules
        elif "lieu de collecte" in line_lower:
            current_obj["Lieu_Collecte"] = extract_value(line)
            current_section = "Lieu_Collecte"
        elif "collecteur" in line_lower:
            current_obj["Collecteur"] = extract_value(line)
            current_section = "Collecteur"
        elif "date de collecte" in line_lower:
            current_obj["Date_Collecte"] = extract_value(line)
            current_section = "Date_Collecte"
            
        # Inscriptions / Transcription Rules
        elif "transcription :" in line_lower:
            val = extract_value(line)
            current_obj["Inscriptions"].append(val if val else "")
            current_section = "Transcription"
            
        elif "référence bibliographique" in line_lower or "reference bibliographique" in line_lower:
            current_obj["Bibliography"] = extract_value(line)
            current_section = None
            
        elif "notes :" in line_lower:
            val = extract_value(line)
            if current_block == "bibliography":
                current_obj["Biblio_Notes"] = val
                current_section = None
            else:
                if val:
                    current_obj["General_Notes"].append(val)
                current_section = "General_Notes"
    
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
                    if current_obj["Function_Role"]:
                        current_obj["Function_Role"] += ", " + line
                    else:
                        current_obj["Function_Role"] = line
                elif current_section == "Date_Fabrication":
                    if current_obj["Date_Fabrication"]:
                        current_obj["Date_Fabrication"] += " " + line
                    else:
                        current_obj["Date_Fabrication"] = line
                elif current_section == "Lieu_Fabrication":
                    if current_obj["Lieu_Fabrication"]:
                        current_obj["Lieu_Fabrication"] += " " + line
                    else:
                        current_obj["Lieu_Fabrication"] = line
                elif current_section == "Date_Utilisation":
                    if current_obj["Date_Utilisation"]:
                        current_obj["Date_Utilisation"] += " " + line
                    else:
                        current_obj["Date_Utilisation"] = line
                elif current_section == "Lieu_Utilisation":
                    if current_obj["Lieu_Utilisation"]:
                        current_obj["Lieu_Utilisation"] += " " + line
                    else:
                        current_obj["Lieu_Utilisation"] = line
                elif current_section == "Collecteur":
                    if current_obj["Collecteur"]:
                        current_obj["Collecteur"] += " " + line
                    else:
                        current_obj["Collecteur"] = line
                elif current_section == "Transcription":
                    if current_obj["Inscriptions"]:
                        current_obj["Inscriptions"][-1] += " " + line
                    else:
                        current_obj["Inscriptions"].append(line)
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
        
        # Merge Bibliography with its associated note (e.g. P.786)
        biblio_complete = obj["Bibliography"]
        if obj["Biblio_Notes"]:
            biblio_complete += f" p. {obj['Biblio_Notes']}" if biblio_complete else obj["Biblio_Notes"]
        biblio_complete = clean_for_csv(biblio_complete)
        
        # Concatenate all accumulated general notes cleanly
        # E.g. "Note 1 text. Dimensions de la 2e partie : Hauteur : 36"
        general_notes_clean = clean_for_csv(" --- ".join(obj["General_Notes"]))
        
        # Format Inscriptions separated by " / "
        inscriptions_joined = clean_for_csv(" / ".join([i for i in obj["Inscriptions"] if i.strip()]))
        
        row = [
            clean_for_csv(obj["Inventory"]),
            desig1,
            desig2,
            clean_for_csv(obj["Function_Role"]),
            "", # Fabricant (no pattern defined yet)
            clean_for_csv(obj["Lieu_Fabrication"]),
            clean_for_csv(obj["Date_Fabrication"]),
            clean_for_csv(obj["Date_Utilisation"]),
            clean_for_csv(obj["Lieu_Utilisation"]),
            materials_joined,
            clean_for_csv(obj["Height"]),    
            clean_for_csv(obj["Width"]),     
            clean_for_csv(obj["Thickness"]), 
            description_clean,
            domain_joined,
            clean_for_csv(obj["Acquisition"]),
            clean_for_csv(obj["Date_Collecte"]),
            clean_for_csv(obj["Lieu_Collecte"]),
            "", # Donateur (no pattern defined yet)
            clean_for_csv(obj["Collecteur"]),
            inscriptions_joined,
            "", # Description restauration (no pattern defined yet)
            "", # Date de restauration (no pattern defined yet)
            biblio_complete,
            general_notes_clean  
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