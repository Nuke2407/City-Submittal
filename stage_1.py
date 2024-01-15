import pandas as pd
import numpy as np

class Stage_1():
    def SQs_missing_data(self):
        self.import_manupulate_data()
        #Create conditions for filtering rows.
        conditions = self.SQ_metadata_closed_file[['Design Directive', 'Category of Change', 'Class of Change']].replace('', np.nan).isna()
        self.missing_metadata_file = self.SQ_metadata_closed_file[conditions.any(axis=1)]
        #Select only the specific columns for self.missing_metadata_file.
        self.missing_metadata_file = self.missing_metadata_file[['Document No', 'Title', 'Discipline', 'Category of Change', 'Design Directive', 'Class of Change']]

        #Creates a new dataframe where the Design Package, Incorporated To, and Specifications are split up and converted into lists for better comparison.
        self.incorrect_metadata_file = self.SQ_metadata_closed_file
        pattern_packages = r'(VLW-PKG-.*?)(?=, VLW-PKG-|$)'
        pattern_specs = r'(VLW-SPC-.*?)(?=, VLW-SPC-|$)'
        pattern_combined = r'(VLW-(?:PKG|SPC)-.*?)(?=, VLW-(?:PKG|SPC)-|$)'
        self.incorrect_metadata_file['Design Package Split Up'] = self.incorrect_metadata_file['Design Package'].str.extractall(pattern_packages)[0].groupby(level=0).apply(list)
        self.incorrect_metadata_file['Specs Split Up'] = self.incorrect_metadata_file['Specifications'].str.extractall(pattern_specs)[0].groupby(level=0).apply(list)
        self.incorrect_metadata_file['Incorporated Design Package Split Up'] = self.incorrect_metadata_file['Incorporated To'].str.extractall(pattern_combined)[0].groupby(level=0).apply(list)


        self.incorrect_metadata_file = Stage_1.incorrect_metadata_check(self.incorrect_metadata_file)
        self.incorrect_metadata_file = self.incorrect_metadata_file[['Document No', 'Title', 'Design Directive', 'Class of Change', 'Category of Change', 
                                                                     'Reason for Change', 'Design Package', 'Specifications', 'Incorporated To', 'Error Type']]

        #Create a downloadable version of the dataframe that contains both the incorrect and missing metadata. 
        with pd.ExcelWriter('data/missing_incorrect_metadata_file.xlsx') as writer:
            self.missing_metadata_file.to_excel(writer, sheet_name='SQs With Missing Data', index=False)
            self.incorrect_metadata_file.to_excel(writer, sheet_name='SQs With Incorrect Data', index=False)

        # Replace NaN values with an empty string.
        self.missing_metadata_file.fillna('', inplace=True)
        
        #return self.missing_metadata_file

    def import_manupulate_data(self):
        "Imports all the metadata formatting it in order to output the documents missing data."
        excel_file_1 = 'stage_one_documents/SQ_metadata_closed_template.xlsx'
        excel_file_2 = 'stage_one_documents/ExportDocs_Stage_1.xls'
        excel_file_3 = 'stage_one_documents/VLW-LOG-11000050-DC-0001_SQ_old.xls'

        #Reads the excel file and stores the datat in Dataframes. The 'header=' sets the header to the 12th excel column.
        self.SQ_metadata_closed_file = pd.read_excel(excel_file_1, header=0)
        ExportDocs = pd.read_excel(excel_file_2, header=10)
        VLW_LOG_11000050_DC_0001_SQ_old = pd.read_excel(excel_file_3)

        #Copying data from ExportDocs to self.SQ_metadata_closed_file.
        data_to_check = ExportDocs.loc[0:, 'Document No':'Date Reviewed']
 
        #Check to see how many rows have to be added.
        required_rows = len(data_to_check)
        current_rows = len(self.SQ_metadata_closed_file)
        additional_rows_needed = required_rows - current_rows

        #If there are not enough rows adds some so that the data fits.
        if additional_rows_needed > 0:
            additional_df = pd.DataFrame(np.nan, index=range(additional_rows_needed), columns=self.SQ_metadata_closed_file.columns)
            self.SQ_metadata_closed_file = pd.concat([self.SQ_metadata_closed_file, additional_df], ignore_index=True)

        #Copies the data over from the Export document to the SQ_metadata_closed_file.
        for (cl_name, data_series) in data_to_check.items():
            self.SQ_metadata_closed_file[cl_name] = data_series 

        #Create a dictionary of document numbers and their revision numbers from VLW_LOG_11000050_DC_0001_SQ_old.
        vlw_doc_and_revisions = {row['Document No']: row['Revision'] for index, row in VLW_LOG_11000050_DC_0001_SQ_old.iterrows()}
        #Define a function to apply to each row in self.SQ_metadata_closed_file.
        def check_SQ_Rev(row):
            doc_number = row['Document No']
            current_revision = row['Revision']
            if doc_number in vlw_doc_and_revisions:
                return 'Y' if current_revision > vlw_doc_and_revisions[doc_number] else 'N'
            else:
                return 'N/A'
            
        #Apply the function to self.SQ_metadata_closed_file.
        self.SQ_metadata_closed_file['Rev. updated?(Y/N/NA)'] = self.SQ_metadata_closed_file.apply(check_SQ_Rev, axis=1)
        self.SQ_metadata_closed_file['If already sent to City?(Y/N)'] = self.SQ_metadata_closed_file['Document No'].apply(lambda x: 'Y' if x in vlw_doc_and_revisions else 'N')

        #Creates an excel file without the index on the left side.
        self.SQ_metadata_closed_file.to_excel('data/SQ_metadata_closed_file.xlsx', index=False)
    
    def incorrect_metadata_check(incorrect_metadata_file):
        incorrect_row = []
        
        #Iterating over the entire dataframe to check the if the metadata has been fillout correctly.
        for index, row in incorrect_metadata_file.iterrows():
            #Checking completeness of the metadata associated with the Class of Change. 
            if row['Class of Change'] == 'Incorporated in Design':
                package_spec_combined_set = set()
                incorporated_set = set() 

                if isinstance(row['Design Package Split Up'], list):
                    package_spec_combined_set.update(row['Design Package Split Up'])
                if isinstance(row['Specs Split Up'], list):
                    package_spec_combined_set.update(row['Specs Split Up'])
                if isinstance(row['Incorporated Design Package Split Up'], list):
                    incorporated_set.update(row['Incorporated Design Package Split Up'])

                if package_spec_combined_set != incorporated_set:
                    row['Error Type'] = 'Incorrect Packages/Specs for INCORPORATED in Design class of change'
                    incorrect_row.append(row)

            elif row['Class of Change'] == 'Partially Incorporated in Design':
                
                incorporated_empty = not isinstance(row['Incorporated Design Package Split Up'], list)
                package_spec_empty = not isinstance(row['Specs Split Up'], list) and not isinstance(row['Design Package Split Up'], list)

                package_spec_combined_set = set(row['Design Package Split Up']) if isinstance(row['Design Package Split Up'],list) else set()
                package_spec_combined_set.update(set(row['Specs Split Up']) if isinstance(row['Specs Split Up'],list) else set())
                incorporated_set = set(row['Incorporated Design Package Split Up']) if isinstance(row['Incorporated Design Package Split Up'],list) else set()

                sets_intersect = not package_spec_combined_set.isdisjoint(incorporated_set)

                if incorporated_empty or package_spec_empty or sets_intersect:
                    row['Error Type'] = 'Incorrect Packages/Specs for PARTIALLY INCORPORATED in Design class of change'
                    incorrect_row.append(row)

            elif row['Class of Change'] == 'Not Incorporated in Design':

                incorporated_empty = isinstance(row['Incorporated Design Package Split Up'], list)
                package_spec_empty = not isinstance(row['Specs Split Up'], list) and not isinstance(row['Design Package Split Up'], list)

                if incorporated_empty or package_spec_empty:
                    row['Error Type'] = 'Incorrect Packages/Specs for NOT INCORPORATED in Design class of change'
                    incorrect_row.append(row)

            elif row['Class of Change'] == 'Field Redline':
#May be a source of error later on double check with the team.
                incorporated_empty = isinstance(row['Incorporated Design Package Split Up'], list)
                spec_empty = isinstance(row['Specs Split Up'], list)

                if incorporated_empty or spec_empty:
                    row['Error Type'] = 'Incorrect Packages/Specs for Field Redline class of change'
                    incorrect_row.append(row)

            #Need to figure out how to check for both 
            elif row['Class of Change'] == 'Field Redline, Incorporated in Design':
                None
            elif row['Class of Change'] == 'Field Redline, Not Incorporated in Design':    
                None

            #Checking completeness of the metadata associated with the Design Directive. 
            if row['Design Directive'] == 'RFI – No Design Change':
                class_of_change_check = not row['Class of Change'] == 'No Action Required' 
                reason_for_change_check = not row['Reason for Change'] == 'Request for Information/Clarification' 
                category_of_change_check = not row['Category of Change'] == 'No Change (response does not result in design change)' 

                if class_of_change_check or reason_for_change_check or category_of_change_check:
                    row['Error Type'] = 'Metadata associated with RFI Design Directive is Incorrect'
                    incorrect_row.append(row)

            elif row['Design Directive'] == 'FCD – Field Change Directive':
                class_of_change_check = row['Class of Change'] == 'No Action Required' or row['Class of Change'] == 'Non-Design Change' or row['Class of Change'] in [None, '', np.nan]
                reason_for_change_check = row['Reason for Change'] == 'Request for Information/Clarification' or row['Reason for Change'] in [None, '', np.nan]
                category_of_change_check = row['Category of Change'] == 'No Change (response does not result in design change)' or row['Category of Change'] in [None, '', np.nan]

                if class_of_change_check or reason_for_change_check or category_of_change_check:
                    row['Error Type'] = 'Metadata associated with FCD Design Directive is Incorrect'
                    incorrect_row.append(row)

        return pd.DataFrame(incorrect_row)
    
    def edit_missing_metadata_file(self): 
        # Create a Pandas Excel writer using XlsxWriter as the engine.
        writer = pd.ExcelWriter('data/missing_incorrect_metadata_file.xlsx', engine='xlsxwriter')
        # Write the dataframe data to XlsxWriter
        self.missing_metadata_file.to_excel(writer, sheet_name='Sheet1', index=False)

         # Get the xlsxwriter workbook and worksheet objects
        workbook  = writer.book
        worksheet = writer.sheets['Sheet1']

        # Define a format for the header
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': '#0EC5D8',
            'border': 1})
        
        # Define a format for the data cells with borders
        cell_format = workbook.add_format({'border': 1})  # This line defines cell_format
        
        # Write the column headers with the defined format
        for col_num, value in enumerate(self.missing_metadata_file.columns.values):
            worksheet.write(0, col_num, value, header_format)

        # Find the maximum length of the content in each column
            column_len = self.missing_metadata_file[value].astype(str).map(len).max()
            column_len = max(column_len, len(value))  # compare with column header length
            worksheet.set_column(col_num, col_num, column_len)  # set column width

         # Apply the cell format to each cell in the data
        for row_num in range(1, len(self.missing_metadata_file) + 1):
            for col_num in range(len(self.missing_metadata_file.columns)):
                worksheet.write(row_num, col_num, self.missing_metadata_file.iloc[row_num - 1, col_num], cell_format)

        # Close the Pandas Excel writer and output the Excel file
        writer.close()

    