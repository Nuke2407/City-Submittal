import pandas as pd
import numpy as np

class Stage_1():
    def importdata():
        "Imports all the metadata formatting it in order to output the documents missing data."
        excel_file_1 = 'stage_one_documents/SQ_metadata_closed_template.xlsx'
        excel_file_2 = 'stage_one_documents/ExportDocs.xls'
        excel_file_3 = 'stage_one_documents/VLW-LOG-11000050-DC-0001_SQ_old.xls'

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
        SQ_metadata_closed_file_columns = SQ_metadata_closed_file.columns[3:15]

        #Copies the data over 
        for dest_col, (cl_name, data_series) in zip(SQ_metadata_closed_file_columns, data_to_check.items()):
            SQ_metadata_closed_file[dest_col] = data_series


        #Create a dictionary of document numbers and their revision numbers from VLW_LOG_11000050_DC_0001_SQ_old.
        vlw_doc_and_revisions = {row['Document No']: row['Revision'] for index, row in VLW_LOG_11000050_DC_0001_SQ_old.iterrows()}
        #Define a function to apply to each row in SQ_metadata_closed_file
        def check_document_status(row):
            doc_number = row['Document No']
            current_revision = row['Revision']
            if doc_number in vlw_doc_and_revisions:
                return 'Y' if current_revision > vlw_doc_and_revisions[doc_number] else 'N'
            else:
                return 'N/A'
        #Apply the function to SQ_metadata_closed_file
        SQ_metadata_closed_file['Rev. updated?(Y/N/NA)'] = SQ_metadata_closed_file.apply(check_document_status, axis=1)
        SQ_metadata_closed_file['If already sent to City?(Y/N)'] = SQ_metadata_closed_file['Document No'].apply(lambda x: 'Y' if x in vlw_doc_and_revisions else 'N')

        #Create a dataframe to display the metadata that has not been updated. 
        conditions = (SQ_metadata_closed_file[['Design Directive', 'Category of Change', 'Class of Change']].replace('', np.nan).isna())
        filtered_rows = SQ_metadata_closed_file[conditions.any(axis=1)]

        print(filtered_rows)

        #Creates an excel file without the index on the left side. 
        SQ_metadata_closed_file.to_excel('data/SQ_metadata_closed_file.xlsx', index=False)

Stage_1.importdata()

