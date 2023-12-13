import os
import pandas as pd
from glob import glob

#%% 

# Replace 'C:/test_excel' with the actual path to your folder
folder_path = 'C:/test_excel'

# Get a list of all Excel files in the folder with both .xlsx and .xlt extensions
excel_files = glob(os.path.join(folder_path, '*.[xX][lL][sS][xX]')) + glob(os.path.join(folder_path, '*.[xX][lL][tT]'))

# Convert \ to / 
excel_files = [file_path.replace('\\', '/') for file_path in excel_files]

#%%
for file_path in excel_files:
    print(file_path)
    df = pd.read_excel(file_path) 
    

    
    # Extract the first two words from the second row
    first_two_words = df.iloc[1, 0].split()[:2]
    

    # Remove the last character from the second word
    if len(first_two_words) == 2:
        first_two_words[1] = first_two_words[1][:-1]

    # Join the words to create the name for the new Excel file
    new_excel_name = ' '.join(first_two_words)

    # Read the Excel file, skipping the first few rows since they contain header information
    # If your decimals use . remove the decimal=','
    df = pd.read_excel(file_path, skiprows=4 ,  decimal=',')
    
    # Change the column names to descending order
    df.columns = range(1, len(df.columns) + 1)[::-1]
    
    # Find the 1st and last Non Null value to extract 
    def extract_numeric_values(condition):
        start_index = df[df.iloc[:, 0] == condition].index[0]
        numeric_values = df.loc[start_index, :].apply(pd.to_numeric, errors='coerce')
        return numeric_values[~numeric_values.isna()]
    
 
    # The values you want to extract 
    names = ['Γλυκόζη' , 'Ουρία' , 'Κρεατινίνη' , 'SGOT (AST)']
   
    names_numeric_values = {}
    
    for name in names:
       names_numeric_values[name] = extract_numeric_values(name)

    result_df = pd.DataFrame({name: [names_numeric_values[name].iloc[0], names_numeric_values[name].iloc[-1]] for name in names})
    
    # Save the DataFrame to a new Excel file
    output_excel_path =os.path.join(folder_path, f'{new_excel_name}.xlsx')
    result_df.to_excel(output_excel_path, index=False)
    print('extracted values in ' , output_excel_path)