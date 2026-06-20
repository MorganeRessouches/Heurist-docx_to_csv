# Document Inventory Parser

This is a Streamlit web application designed to parse structured catalog/inventory tables from Microsoft Word (.docx) documents and export the parsed information into a clean, standardized CSV file formatted with semicolon (`;`) separators.

The parsing engine analyzes text blocks, identifies keys (optimized for French museum/inventory terminology), groups fields dynamically, and processes multi-line blocks like descriptions and materials list.

## Features

- **File Upload**: Drag and drop any `.docx` file containing data tables.
- **In-Memory Parsing**: Document processing is conducted entirely in memory, without local disk write requirements.
- **Verification Logs**: Identifies tables that do not contain valid inventory metadata tags to simplify structure verification.
- **Data Preview**: View the first few parsed records in a clear tabular interface before downloading.
- **Custom Mapping**: Standardizes inventory numbers, item designations, dimensions, and bibliographic records into formatted columns.

## Project Structure

```text
├── app.py             # Streamlit application entrypoint
├── requirements.txt   # Python dependency specifications
└── README.md          # Project documentation
```

## Installation

### 1. Prerequisites
Ensure you have Python 3.8 or higher installed on your system.

### 2. Clone or Download the Project
Download the source files into a local directory of your choice.

### 3. Install Required Dependencies
Open a terminal in the root directory of the project and run:

```bash
pip install -r requirements.txt
```

The application relies on the following packages:
* `streamlit` - For the interactive web interface.
* `python-docx` - To extract elements from MS Word tables.
* `pandas` - For dataframe representation and data conversion.

## How to Run

To run the application locally, execute the following command in your terminal:

```bash
streamlit run app.py
```

A local development server will start, and the application will automatically open in your default browser (typically at `http://localhost:8501`).

## Data Conversion Logic Details

- **Object Separation**: A new record is registered when the parser encounters a line beginning with `"numéro d'inventaire"`.
- **Dimensions Mapping**:
  - `Hauteur en cm` maps to the **Length** CSV field.
  - `Largeur en cm` maps to the **Width** CSV field.
  - `Epaisseur en cm` maps to the **Height** CSV field.
- **Bibliography and Notes**: If footnotes or page values are present under the `Notes` field, they are appended to the `Bibliographical References` column as `[Bibliography] p. [Notes]`.
```