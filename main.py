'''
Script to find the excel files from the same directory and checks the column names and then
creates a new excel file with the same columns merged including different colunms.
'''
import os
import pandas as pd
import datetime

def get_files() -> list:
    '''
    Function to get the excel files from the folders in same dir and returns the list of files
    Check for files in the same directory and then search folders for the excel files
    '''
    files = [file for file in os.listdir() if file.endswith('.xlsx')]
    if not files:
        for folder in os.listdir():
            if os.path.isdir(folder):
                # ignore the files with startswith mergerd
                files.extend([f'{folder}/{file}' for file in os.listdir(folder) if file.endswith('.xlsx') and not file.startswith('merged')])
    return files
    

def get_columns(files) -> list:
    '''
    Function to get the columns from the excel files and returns the list of columns
    '''
    columns = []
    for file in files:
        df = pd.read_excel(file)
        columns.append(df.columns)
    return columns

def get_common_columns(columns) -> list:
    '''
    Function to get the common columns from the excel files and returns the list of common columns
    '''
    common_columns = set(columns[0])
    for column in columns:
        common_columns = common_columns.intersection(set(column))
    return list(common_columns)

def get_unique_columns(columns, common_columns) -> list:
    '''
    Function to get the unique columns from the excel files and returns the list of unique columns
    '''
    unique_columns = []
    for column in columns:
        unique_columns.append(list(set(column) - set(common_columns)))
    return unique_columns

def get_dataframes(files, common_columns, unique_columns) -> list:
    '''
    Function to get the dataframes from the excel files
    '''
    dataframes = []
    for index, file in enumerate(files):
        df = pd.read_excel(file)
        dataframes.append(df[common_columns + unique_columns[index]])
    return dataframes

def merge_dataframes(dataframes) -> pd.DataFrame:
    '''
    Function to merge the dataframes with sorted columns
    '''
    merged_df = pd.concat(dataframes, sort=True)
    return merged_df

def write_to_excel(merged_df) -> None:
    '''
    Function to remove empty cols and then write the merged dataframe to excel
    '''
    merged_df = merged_df.dropna(axis=1, how='all')
    # Rename the file merged+current date and time
    merged_df.to_excel(f'merged_{datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")}.xlsx', index=False)


def main():
    '''
    Main function
    '''
    files = get_files()
    columns = get_columns(files)
    common_columns = get_common_columns(columns)
    unique_columns = get_unique_columns(columns, common_columns)
    dataframes = get_dataframes(files, common_columns, unique_columns)
    merged_df = merge_dataframes(dataframes)
    write_to_excel(merged_df)

if __name__ == '__main__':
    main()
    