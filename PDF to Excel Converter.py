#pip install pdfplumber pandas openpyxl


import pdfplumber
import pandas as pd

def pdf_to_excel(pdf_file, excel_file):
    # Create a list to hold dataframes
    dataframes = []

    # Open the PDF file
    with pdfplumber.open(pdf_file) as pdf:
        # Loop through each page of the PDF
        for page in pdf.pages:
            # Extract tables from the page
            tables = page.extract_tables()
            for table in tables:
                # Convert the table to a DataFrame and append to the list
                df = pd.DataFrame(table[1:], columns=table[0])
                dataframes.append(df)

    # Concatenate all DataFrames into a single DataFrame
    final_df = pd.concat(dataframes, ignore_index=True)

    # Save the final DataFrame to an Excel file
    final_df.to_excel(excel_file, index=False)

    print(f"Converted {pdf_file} to {excel_file}")

if __name__ == "__main__":
    pdf_file_path = "C:/PDF to Excel Converter/BSTDB Financial Statements for 2021.pdf"  # Specify your PDF file path here
    excel_file_path = "C:/Users/Viswajith/Downloads/output1.xlsx"          # Specify the output Excel file name here
    pdf_to_excel(pdf_file_path, excel_file_path)









# import pdfplumber
# import pandas as pd

# def pdf_to_excel(pdf_file, excel_file):
#     # Create a list to hold dataframes
#     dataframes = []

#     # Open the PDF file
#     with pdfplumber.open(pdf_file) as pdf:
#         # Loop through each page of the PDF
#         for page in pdf.pages:
#             # Extract tables from the page
#             tables = page.extract_tables()
#             for table in tables:
#                 # Convert the table to a DataFrame
#                 df = pd.DataFrame(table[1:], columns=table[0])
                
#                 # Ensure unique column names
#                 df.columns = [f"{col}_{i}" if df.columns.tolist().count(col) > 1 else col for i, col in enumerate(df.columns)]
                
#                 dataframes.append(df)

#     # Concatenate all DataFrames into a single DataFrame
#     final_df = pd.concat(dataframes, ignore_index=True)

#     # Save the final DataFrame to an Excel file
#     final_df.to_excel(excel_file, index=False)

#     print(f"Converted {pdf_file} to {excel_file}")

# if __name__ == "__main__":
#     pdf_file_path = "C:/Users/Viswajith/Downloads/wipo_pub_finstatements_2016.pdf"  # Specify your PDF file path here
#     excel_file_path = "C:/Users/Viswajith/Downloads/output1.xlsx"          # Specify the output Excel file name here
#     pdf_to_excel(pdf_file_path, excel_file_path)







