import camelot
import pandas as pd
import re

target_page = 67  # Replace with your desired page number
offset = 2


def custom_table_extraction_settings():
    # Custom settings for table extraction
    settings = {
        'edge_tol': 300,
        'text_tol': 15
    }
    return settings



# Calculate the actual page number
actual_page = target_page + offset


custom_settings = custom_table_extraction_settings()
# Read PDF and extract tables from the specified page
tables = camelot.read_pdf('Data\CAN110056_Electronics_Price_Book_.pdf', pages=str(actual_page), flavor='stream',**custom_settings)


# Function to extract numbers from a string, including characters after the digits
def extract_numbers_with_characters(text):
    return [match.group() for match in re.finditer(r'\b\d+\S*\b', text)]

# List to store the extracted numbers
extracted_numbers = []
concatenated_df = pd.DataFrame()


#df to excel
for i in range(len(tables)):
 tables[i].to_excel('file_'+str(i)+'.xlsx')

#print the df in terminal
for i, table in enumerate(tables):
    print(f"Table {i + 1}:")
    print(table.df)
    print("\n")


    finishes_found = False

    # Iterate through each row in the table
    for index, row in table.df.iterrows():
        # Check if "Available finishes" is in the row
        if "Available finishes" in row.values:
            # Set the flag to True
            finishes_found = True

        # If "Available finishes" has been found and the flag is True, extract numbers from the current row
        elif finishes_found:
            cell_values = row.tolist()
            for cell_value in cell_values:
                numbers = extract_numbers_with_characters(str(cell_value))
                extracted_numbers.extend(numbers)
            # Set the flag to False after processing the row below "Available finishes"
            finishes_found = False
            break

# Display the extracted numbers
print("Extracted Numbers:", extracted_numbers)

# Loop through each table and concatenate the header and rows 6 to 8
for i, table in enumerate(tables):
    header = table.df.iloc[5:7, :]
    part_to_concat = table.df.iloc[6:9, :]
    part_to_concat2 = table.df.iloc[7:9, :]  
    concatenated_df = pd.concat([header])
    for j in range(len(extracted_numbers)): 
        concatenated_table = pd.concat([part_to_concat2], ignore_index=True)
        concatenated_df = pd.concat([concatenated_df, concatenated_table], ignore_index=True)
        concatenated_df=concatenated_df.replace({
            '.[Finish]':extracted_numbers[j]
        })



    # df to excel for each table
    #concatenated_table.to_excel(f'concatenated_file_{i}.xlsx')

# Print the concatenated DataFrame in the terminal
print("Concatenated DataFrame:")
print(concatenated_df)
print("\n")

# Print the concatenated DataFrame as Excel
concatenated_df.to_excel('concatenated_file.xlsx')




