# hs-code-tree-to-excel
# HS Code Exporter

This project provides a Python script to automatically download, parse, and export the latest Harmonized System (HS) code data from the United Nations Statistics Division (UNSD) into a structured Excel file.

## Features
- Downloads the most up-to-date HS code data (HS6 level) from the official UNSD source.
- Parses and organizes the data into a hierarchical structure:
  - **HS2**: Chapters (2-digit)
  - **HS4**: Headings (4-digit)
  - **HS6**: Subheadings (6-digit)
- Exports to Excel with two sheets:
  - **HS Tree**: Tree-like, indented view for easy navigation.
  - **Flat Table**: Simple table with all codes and descriptions.

## Requirements
- Python 3.7+
- Packages: `requests`, `pandas`, `openpyxl`

Install dependencies:
```sh
pip install requests pandas openpyxl
```

## Usage
Run the script from the project directory:
```sh
python nts/hs_code_exporter.py
```

This will create `HS_2022_Codes.xlsx` in your working directory.

## Output
- **HS Tree Sheet**: Shows the HS code hierarchy in a tree format. Each row contains one code at its level, with indentation for HS2, HS4, and HS6.
- **Flat Table Sheet**: Contains all HS6 codes with their descriptions and parent codes for reference.

## Updating the HS Code Source
If the download link changes:
1. Visit [UNSD Classifications on Economic Statistics](https://unstats.un.org/unsd/classifications/Econ/)
2. Find the latest "All HS codes and descriptions" (XLSX) link under the HS section.
3. Update the `HS_URL` variable in `nts/hs_code_exporter.py` with the new link.

## HS Code Structure Explained
- **HS2 (Chapter)**: First 2 digits, broadest category (e.g., `01` = Live animals)
- **HS4 (Heading)**: First 4 digits, more specific (e.g., `0101` = Horses, asses, mules, and hinnies)
- **HS6 (Subheading)**: Full 6 digits, most detailed international level (e.g., `010121` = Pure-bred breeding horses)

## Extending the Script
The script is modular and can be extended for:
- Country-specific codes (HS8/HS10)
- Language localization
- User input for custom exports

---
For questions or contributions, please open an issue or pull request. 
