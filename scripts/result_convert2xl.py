import win32com.client as win32
import pandas as pd
import os

'''please just give your folder location that contains your .vd result files'''
# location that contains .vd files
input_folder_path = "C:/AnswerTIMESv6/Gams_WrkTI/"

# scenario result  files
'''please specify the scenarios results file (.vd) name here'''
result_files = [
    input_folder_path + 'BALANCE.vd',
    input_folder_path + 'ECPRICES_HYDROGE.vd'
]
'''please provide local location of output (result file) and just write result file name (e.g. all_result.txt),
The text file will be generated automatically. 
'''
# local location of output (result file)
#output_file_path = "C:/Users/ac141435/Desktop/TAM_Idustrie/TAM-Industry/results/all_results_ich.txt"
output_file_path = "result_text/all_results.txt"

with open(output_file_path, 'w') as all_result:  # create all_result text file to write data (line) from .vd file
    #for file_name in os.listdir(input_folder_path):
    for file_name in result_files: # loop through each given
        if file_name.endswith('.vd'):  # read only the .vd file from input folder
            input_file_path = os.path.join(input_folder_path, file_name)  # read input file path from input folder
            print(f'Your input file {input_file_path}')
            scenario_name, _ = os.path.splitext(os.path.basename(input_file_path))  # extract scenario name from .vd
            with open(input_file_path, 'r') as file:  # open input .vd as file to read data (line)
                data = file.readlines()  # all lines are stored in data

            for line in data:
                if not line.strip():  # Skip empty or whitespace-only lines
                    continue
                if not line.startswith('*'):
                    if f'{scenario_name},' not in line:
                        # create scenario column
                        all_result.write(f'"{scenario_name}",{line}')  # write scenario name in each row (line)

                    else:
                        all_result.write(line)
                        print('Should not execute ELSE, is something wrong?')

print(f'your output file {all_result.name} is created successfully.')


# local location of all_result.text
data_path = output_file_path

# column names (headers in Excel)
col_names = ['Scenario', 'Attribute', 'Commodity', 'Process', 'Period', 'Region', 'Vintage',
             'TimeSlice', 'UserConstraint', 'PV']

data = []  # empty list to append data from chunk
# read the text file in chunk
for data_chunk in pd.read_csv(data_path, header=None, names=col_names, chunksize=100000, low_memory=False):
    data.append(data_chunk)  # appending data chunk as list in list

df = pd.concat(data)  # convert the list as dataframe


data_list = df.values.tolist()  # Convert dataframe to list of lists
# insert column header again
data_list.insert(0, col_names)

''' Please, provide Excel file location where you want to save your data'''
# location (path) of Excel file where to save data from .text file
excel_file_path = os.path.abspath("result_excel/TIMES_result.xlsx")

excel = win32.Dispatch('Excel.Application')
excel.Visible = True
excel.DisplayAlerts = True
wb = excel.Workbooks.Open(excel_file_path)

# Create new worksheet
# ws = wb.Sheets.Add()
# ws.Name = 'test1'
'''Write worksheet name where you want to save data'''
ws = wb.Sheets('all_results')  # Excel Worksheet where data is saved/stored
# Write data to worksheet
ws.Range(ws.Cells(1, 1), ws.Cells(len(data_list), len(data_list[0]))).Value = data_list

print('Results are successfully created in excel')
