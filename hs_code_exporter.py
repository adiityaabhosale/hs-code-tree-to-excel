import os
import requests
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# NOTE: If this link breaks, visit https://unstats.un.org/unsd/classifications/Econ/ and look for 'All HS codes and descriptions' (XLSX) under the HS section.
HS_URL = "https://unstats.un.org/unsd/classifications/Econ/download/In%20Text/HSCodeandDescription.xlsx"
# If the above link fails, update it with the latest direct XLSX link from the UNSD site.
OUTPUT_FILE = "HS_2022_Codes.xlsx"


def download_hs_excel(url=HS_URL):
    """
    Download the HS code Excel file from the UN Statistics Division.
    Returns:
        BytesIO: In-memory binary stream of the Excel file.
    """
    response = requests.get(url)
    response.raise_for_status()
    return BytesIO(response.content)


def parse_hs_excel(excel_stream):
    """
    Parse the HS Excel file and extract the hierarchy from the new format.
    Args:
        excel_stream (BytesIO): The Excel file stream.
    Returns:
        DataFrame: Parsed HS code data with columns: Section, HS2, HS4, HS6, Description
    """
    df = pd.read_excel(excel_stream, dtype=str, header=0)
    df.columns = [c.strip() for c in df.columns]
    # Only keep HS codes (Classification == 'H6') and basic levels (IsBasicLevel == '1')
    df = df[(df['Classification'] == 'H6') & (df['IsBasicLevel'] == '1')]
    df = df.fillna("")
    # Prepare columns for HS2, HS4, HS6
    df['HS2'] = df['Code'].apply(lambda x: x[:2] if len(x) >= 2 else "")
    df['HS4'] = df['Code'].apply(lambda x: x[:4] if len(x) >= 4 else "")
    df['HS6'] = df['Code'].apply(lambda x: x[:6] if len(x) >= 6 else "")
    df['Section'] = ""
    df = df[['Section', 'HS2', 'HS4', 'HS6', 'Description']]
    return df


def build_tree_view(df):
    """
    Build a tree-like DataFrame for Excel export.
    Args:
        df (DataFrame): Parsed HS code data.
    Returns:
        DataFrame: Tree-structured DataFrame for Excel export.
    """
    tree_rows = []
    last_section = last_hs2 = last_hs4 = None
    for _, row in df.iterrows():
        section = row['Section']
        hs2 = row['HS2']
        hs4 = row['HS4']
        hs6 = row['HS6']
        desc = row['Description']
        # Add Chapter (HS2) if new
        if hs2 != last_hs2:
            tree_rows.append([hs2, '', '', desc if len(hs2) == 2 else ''])
            last_hs2 = hs2
            last_hs4 = None
        # Add Heading (HS4) if new
        if hs4 != last_hs4 and len(hs4) == 4:
            tree_rows.append(['', hs4, '', desc if len(hs4) == 4 else ''])
            last_hs4 = hs4
        # Always add Subheading (HS6)
        if len(hs6) == 6:
            tree_rows.append(['', '', hs6, desc])
    tree_df = pd.DataFrame(tree_rows, columns=['HS2', 'HS4', 'HS6', 'Description'])
    return tree_df


def export_to_excel(tree_df, flat_df, output_file=OUTPUT_FILE):
    """
    Export the tree and flat DataFrames to an Excel file with two sheets.
    Args:
        tree_df (DataFrame): Tree-structured DataFrame.
        flat_df (DataFrame): Flat DataFrame.
        output_file (str): Output Excel file name.
    """
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        tree_df.to_excel(writer, sheet_name='HS Tree', index=False)
        flat_df[['Section', 'HS6', 'Description']].to_excel(writer, sheet_name='Flat Table', index=False)
    print(f"Exported to {output_file}")


def main():
    """
    Main function to orchestrate the download, parse, and export process.
    """
    print("Downloading HS code data from UN...")
    excel_stream = download_hs_excel()
    print("Parsing HS code data...")
    flat_df = parse_hs_excel(excel_stream)
    print("Building tree view...")
    tree_df = build_tree_view(flat_df)
    print("Exporting to Excel...")
    export_to_excel(tree_df, flat_df)
    print("Done.")


if __name__ == "__main__":
    main() 