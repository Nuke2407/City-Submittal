import pandas as pd
import numpy as np

excel_file_1 = 'SQ_metadata_closed_template.xlsx'
excel_file_2 = 'ExportDocs.xls'
excel_file_3 = 'VLW-LOG-11000050-DC-0001_SQ_old.xls'

#Reads the excel file and stores the datat in Dataframes. The 'header=' sets the header to the 12th excel column.
SQ_metadata_closed_file = pd.read_excel(excel_file_1, header=0)
ExportDocs = pd.read_excel(excel_file_2, header=10)
VLW_LOG_11000050_DC_0001_SQ_old = pd.read_excel(excel_file_3)

#Copying data from ExportDocs to SQ_metadata_closed_file
data_to_check = ExportDocs.loc[0:, 'Document No':'Date Reviewed']

#Add necessary rows to SQ_metadata_closed_file
required_rows = len(data_to_check)
current_rows = len(SQ_metadata_closed_file)
additional_rows_needed = required_rows - current_rows

if additional_rows_needed > 0:
    additional_df = pd.DataFrame(np.nan, index=range(additional_rows_needed), columns=SQ_metadata_closed_file.columns)
    SQ_metadata_closed_file = pd.concat([SQ_metadata_closed_file, additional_df], ignore_index=True)

#Define the columns to be copied to
SQ_metadata_closed_file_columns = SQ_metadata_closed_file.columns[4:16]

#Copies the data over 
for dest_col, (cl_name, data_series) in zip(SQ_metadata_closed_file_columns, data_to_check.items()):
    SQ_metadata_closed_file[dest_col] = data_series

#SQ_metadata_closed_file.loc[1:, 'Metadata Status Update'] = ""
SQ_metadata_closed_file.to_excel('mnt/data/SQ_metadata_closed_file.xlsx', index=False)
