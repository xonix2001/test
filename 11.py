import pandas as pd

# Load the Excel file
df = pd.read_excel('path_to_your_excel_file.xlsx')

# Define the columns to check for empty cells and the columns to search for keywords
columns_to_check = ['column_name_1', 'column_name_2']  # Replace with your actual column names
keyword_columns = ['column_name_3', 'column_name_4']  # Replace with your actual column names
keywords = ['keyword1', 'keyword2']  # Replace with your actual keywords

# Function to search for keywords and retrieve data
def retrieve_data_from_keywords(row, keyword_columns, keywords):
    for column in keyword_columns:
        cell_value = str(row[column])
        for keyword in keywords:
            if keyword in cell_value:
                return cell_value
    return None

# Iterate over each row
for index, row in df.iterrows():
    for column in columns_to_check:
        if pd.isna(row[column]):  # If cell is empty
            # Retrieve data from keywords
            data_to_fill = retrieve_data_from_keywords(row, keyword_columns, keywords)
            if data_to_fill:
                df.at[index, column] = data_to_fill

# Save the modified Excel file
df.to_excel('path_to_your_modified_excel_file.xlsx', index=False)




def extract_sentence_form_cell():
    extract_sentences=[]
    sentences=str(cell.value).split('。')
    for i in sentences:
        if any(keyword in i for keyword in keywords):
            extract_sentences.append(i)
    return extract_sentences

def extract_row():
    