import pandas as pd
import numpy as np

class Excel_Manipulation():
    def stage_1(self):
        self.import_manupulate_data()

        self.missing_metadata_file_1 = Excel_Manipulation.missing_metadata_check(self.SQ_metadata_closed_file)
        self.incorrect_metadata_file_stage_1 = Excel_Manipulation.incorrect_metadata_check(self.SQ_metadata_closed_file)

        #Create a downloadable version of the dataframe that contains both the incorrect and missing metadata. 
        with pd.ExcelWriter('data/missing_incorrect_metadata_file.xlsx') as writer:
            self.missing_metadata_file_1.to_excel(writer, sheet_name='SQs With Missing Data', index=False)
            self.incorrect_metadata_file_stage_1.to_excel(writer, sheet_name='SQs With Incorrect Data', index=False)

        incorrect_missing_df_sheet_names = [(self.missing_metadata_file_1, 'SQs With Missing Data'), (self.incorrect_metadata_file_stage_1, 'SQs With Incorrect Data')]

        #Format the Excel file
        self.format_and_save_excel(incorrect_missing_df_sheet_names, 'data/missing_incorrect_metadata_file.xlsx')

    def import_manupulate_data(self):
        excel_file_1 = 'stage_one_documents/SQ_metadata_closed_template.xlsx'
        excel_file_2 = 'stage_one_documents/VLW-LOG-11000050-DC-0001_SQ_old.xls'
        excel_file_3 = 'stage_one_documents/ExportDocs_Stage_1.xls'

        #Reads the excel file and stores the datat in Dataframes. The 'header=' sets the header to the 12th excel column.
        self.SQ_metadata_closed_file = pd.read_excel(excel_file_1, header=0)
        VLW_LOG_11000050_DC_0001_SQ_old = pd.read_excel(excel_file_2)

        temp_df = pd.read_excel(excel_file_3, header=None)
        exportDocs_1 = Excel_Manipulation.find_header(temp_df, excel_file_3)

        #Copying data from exportDocs to self.SQ_metadata_closed_file.
        data_to_check = exportDocs_1.loc[0:, 'Document No':'Date Reviewed']
 
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

        #Filter the DataFrame to only include rows with 'N' in the 'If already sent to City?(Y/N)' column.
        self.new_SQs_for_city_submittal = self.SQ_metadata_closed_file.loc[
            (self.SQ_metadata_closed_file['If already sent to City?(Y/N)'] == 'N') & 
            (((self.SQ_metadata_closed_file['Design Directive'] == 'FCD – Field Change Directive') | 
            (self.SQ_metadata_closed_file['Design Directive'].isnull())) | 
            ((self.SQ_metadata_closed_file['Design Directive'] == 'RFI – No Design Change') & 
            ((self.SQ_metadata_closed_file['Category of Change'].isnull()) | 
            (self.SQ_metadata_closed_file['Class of Change'].isnull()))))]
            
        #Filter the DataFrame to only include rows with 'Y' in the 'If already sent to City?(Y/N)' column and rows with 'Y' in the 'Rev. updated?(Y/N/NA)' column.
        self.new_reved_up_SQs_for_city_submittal = self.SQ_metadata_closed_file.loc[(self.SQ_metadata_closed_file['If already sent to City?(Y/N)'] == 'Y') & (self.SQ_metadata_closed_file['Rev. updated?(Y/N/NA)'] == 'Y') & (self.SQ_metadata_closed_file['Design Directive'] == 'FCD – Field Change Directive')]

        with pd.ExcelWriter('data/new_city_sub_SQs.xlsx') as writer:
            self.new_SQs_for_city_submittal.to_excel(writer, sheet_name='New SQs not sent to the City', index=False)
            self.new_reved_up_SQs_for_city_submittal.to_excel(writer, sheet_name='Reved up SQs', index=False)

        new_SQs_for_city_submittal_df_with_sheet_names = [(self.new_SQs_for_city_submittal, 'New SQs not sent to the City'), (self.new_reved_up_SQs_for_city_submittal, 'Reved up SQs')]

        #Format the Excel file
        self.format_and_save_excel(new_SQs_for_city_submittal_df_with_sheet_names, 'data/new_city_sub_SQs.xlsx')

        #Creates an excel file without the index on the left side.
        self.SQ_metadata_closed_file.to_excel('data/SQ_metadata_closed_file.xlsx', index=False)
    
    def missing_metadata_check(df_file):
        #Create conditions for filtering rows.
        conditions = df_file[['Design Directive', 'Category of Change', 'Class of Change']].replace('', np.nan).isna()
        df_file = df_file[conditions.any(axis=1)]
        #Select only the specific columns for self.missing_metadata_file.
        df_file = df_file[['Document No', 'Title', 'Discipline', 'Category of Change', 'Design Directive', 'Class of Change']]

        return df_file

    def incorrect_metadata_check(df_file):
        #Creates a new dataframe where the Design Package, Incorporated To, and Specifications are split up and converted into lists for better comparison.
        incorrect_metadata_file = df_file.copy()
        pattern_packages = r'(VLW-PKG-.*?)(?=, VLW-PKG-|$)'
        pattern_specs = r'(VLW-SPC-.*?)(?=, VLW-SPC-|$)'
        pattern_combined = r'(VLW-(?:PKG|SPC)-.*?)(?=, VLW-(?:PKG|SPC)-|$)'
        incorrect_metadata_file['Design Package Split Up'] = incorrect_metadata_file['Design Package'].str.extractall(pattern_packages)[0].groupby(level=0).apply(list)
        incorrect_metadata_file['Specs Split Up'] = incorrect_metadata_file['Specifications'].str.extractall(pattern_specs)[0].groupby(level=0).apply(list)
        incorrect_metadata_file['Incorporated Design Package Split Up'] = incorrect_metadata_file['Incorporated To'].str.extractall(pattern_combined)[0].groupby(level=0).apply(list)

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
                    new_row = row.copy() 
                    new_row['Error Type'] = 'Incorrect Packages/Specs for INCORPORATED in Design class of change'
                    incorrect_row.append(new_row)

            elif row['Class of Change'] == 'Partially Incorporated in Design':
                
                incorporated_empty = not isinstance(row['Incorporated Design Package Split Up'], list)
                package_spec_empty = not isinstance(row['Specs Split Up'], list) and not isinstance(row['Design Package Split Up'], list)

                package_spec_combined_set = set(row['Design Package Split Up']) if isinstance(row['Design Package Split Up'],list) else set()
                package_spec_combined_set.update(set(row['Specs Split Up']) if isinstance(row['Specs Split Up'],list) else set())
                incorporated_set = set(row['Incorporated Design Package Split Up']) if isinstance(row['Incorporated Design Package Split Up'],list) else set()

                sets_intersect = not package_spec_combined_set.isdisjoint(incorporated_set)

                if incorporated_empty or package_spec_empty or sets_intersect:
                    new_row = row.copy() 
                    new_row['Error Type'] = 'Incorrect Packages/Specs for PARTIALLY INCORPORATED in Design class of change'
                    incorrect_row.append(new_row)

            elif row['Class of Change'] == 'Not Incorporated in Design':

                incorporated_empty = isinstance(row['Incorporated Design Package Split Up'], list)
                package_spec_empty = not isinstance(row['Specs Split Up'], list) and not isinstance(row['Design Package Split Up'], list)

                if incorporated_empty or package_spec_empty:
                    new_row = row.copy() 
                    new_row['Error Type'] = 'Incorrect Packages/Specs for NOT INCORPORATED in Design class of change'
                    incorrect_row.append(new_row)

            elif row['Class of Change'] == 'Field Redline':
#May be a source of error later on double check with the team.
                incorporated_empty = isinstance(row['Incorporated Design Package Split Up'], list)
                spec_empty = isinstance(row['Specs Split Up'], list)

                if incorporated_empty or spec_empty:
                    new_row = row.copy() 
                    new_row['Error Type'] = 'Incorrect Packages/Specs for Field Redline class of change'
                    incorrect_row.append(new_row)

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
                    new_row = row.copy() 
                    new_row['Error Type'] = 'Metadata associated with RFI Design Directive is Incorrect'
                    incorrect_row.append(new_row)

            elif row['Design Directive'] == 'FCD – Field Change Directive':
                class_of_change_check = row['Class of Change'] == 'No Action Required' or row['Class of Change'] == 'Non-Design Change' or row['Class of Change'] in [None, '', np.nan]
                reason_for_change_check = row['Reason for Change'] == 'Request for Information/Clarification' or row['Reason for Change'] in [None, '', np.nan]
                category_of_change_check = row['Category of Change'] == 'No Change (response does not result in design change)' or row['Category of Change'] in [None, '', np.nan]

                if class_of_change_check or reason_for_change_check or category_of_change_check:
                    new_row = row.copy() 
                    new_row['Error Type'] = 'Metadata associated with FCD Design Directive is Incorrect'
                    incorrect_row.append(new_row)

            if row['Discipline'] == 'Design and Construction':
                new_row = row.copy() 
                new_row['Error Type'] = 'Discipline is set to Design and Construction'
                incorrect_row.append(new_row)

        incorrect_metadata_file = pd.DataFrame(incorrect_row)
        incorrect_metadata_file = incorrect_metadata_file[['Document No', 'Title', 'Design Directive', 'Class of Change', 'Category of Change', 
                                                            'Reason for Change', 'Design Package', 'Specifications', 'Incorporated To', 'Error Type']]
        
        return incorrect_metadata_file
    
    def format_and_save_excel(self, dataframes_with_sheet_names, file_path):
        # Create a Pandas Excel writer using XlsxWriter as the engine
        with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
            # Access the workbook and configure it
            workbook = writer.book
            workbook.nan_inf_to_errors = True

            # Define a format for the header
            header_format = workbook.add_format({
                'bold': True,
                'text_wrap': True,
                'valign': 'top',
                'fg_color': '#0EC5D8',
                'border': 1
            })

            # Define a format for the data cells with borders
            cell_format = workbook.add_format({'border': 1})

            # Function to apply formatting to a worksheet
            def format_worksheet(worksheet, dataframe):
                # Write the column headers with the defined format
                for col_num, value in enumerate(dataframe.columns.values):
                    worksheet.write(0, col_num, value, header_format)
                    # Set column width based on the maximum length of the content in each column
                    column_len = max(dataframe[value].astype(str).map(len).max(), len(value))
                    worksheet.set_column(col_num, col_num, column_len)

                # Apply the cell format to each cell in the data
                for row_num in range(1, len(dataframe) + 1):
                    for col_num in range(len(dataframe.columns)):
                        cell_value = dataframe.iloc[row_num - 1, col_num]
                        # Check for NaT values and replace with None
                        if pd.isna(cell_value):
                            cell_value = None
                        worksheet.write(row_num, col_num, cell_value, cell_format)

            # Write each DataFrame to its respective sheet and apply formatting
            for dataframe, sheet_name in dataframes_with_sheet_names:
                # Write the dataframe data to XlsxWriter
                dataframe.to_excel(writer, sheet_name=sheet_name, index=False)
                # Apply formatting to each worksheet
                format_worksheet(writer.sheets[sheet_name], dataframe)
        
    def stage_2(self):
        excel_file_4 = 'stage_one_documents/ExportDocs_Stage_2.xls' 
        temp_df_2 = pd.read_excel(excel_file_4, header=None)
        exportDocs_2 = Excel_Manipulation.find_header(temp_df_2, excel_file_4)

        excel_file_2 = 'stage_one_documents/ExportDocs_Stage_1.xls' 
        temp_df = pd.read_excel(excel_file_2, header=None)
        exportDocs_1 = Excel_Manipulation.find_header(temp_df, excel_file_2)

        excel_file_2 = 'stage_one_documents/VLW-LOG-11000050-DC-0001_SQ_old.xls'
        VLW_LOG_11000050_DC_0001_SQ_old = pd.read_excel(excel_file_2)

        #Create a set of unique identifiers from exportDocs_1.
        unique_identifiers_new_SQs = set(exportDocs_1['Document No'])
        #Filter out rows in exportDocs_2 that are not in exportDocs_1.
        exportDocs_2_filtered = exportDocs_2[exportDocs_2['Document No'].isin(unique_identifiers_new_SQs)]

        #Create a set of unique identifiers from exportDocs_2.
        unique_identifiers_opended_SQs = set(exportDocs_2['Document No'])
        exportDocs_1_filtered = exportDocs_1[~exportDocs_1['Document No'].isin(unique_identifiers_opended_SQs)]
        unique_identifiers_rev_0_open_again = set(VLW_LOG_11000050_DC_0001_SQ_old ['Document No'])
        exportDocs_1_filtered_to_add = exportDocs_1_filtered[exportDocs_1_filtered['Document No'].isin(unique_identifiers_rev_0_open_again)]
        SQ_metadata_closed_df_stage_2 = pd.concat([exportDocs_2_filtered, exportDocs_1_filtered_to_add])
        SQ_metadata_closed_df_stage_2_sorted = SQ_metadata_closed_df_stage_2.sort_values(by= 'Document No')

        SQ_metadata_closed_df_stage_2_sorted.to_excel('data/SQ_metadata_closed_file_2.xlsx', index=False)

        missing_metadata_file_2 = Excel_Manipulation.missing_metadata_check(SQ_metadata_closed_df_stage_2_sorted)
        incorrect_metadata_file_stage_2 = Excel_Manipulation.incorrect_metadata_check(SQ_metadata_closed_df_stage_2_sorted)

        #Create a downloadable version of the dataframe that contains both the incorrect and missing metadata. 
        with pd.ExcelWriter('data/missing_incorrect_metadata_file_2.xlsx') as writer:
            missing_metadata_file_2.to_excel(writer, sheet_name='SQs With Missing Data 2', index=False)
            incorrect_metadata_file_stage_2.to_excel(writer, sheet_name='SQs With Incorrect Data 2', index=False)
            exportDocs_1_filtered_to_add.to_excel(writer, sheet_name='test', index=False)

        incorrect_missing_df_sheet_names_stage_2 = [(missing_metadata_file_2, 'SQs With Missing Data 2'), (incorrect_metadata_file_stage_2, 'SQs With Incorrect Data 2')]

        #Format the Excel file
        self.format_and_save_excel(incorrect_missing_df_sheet_names_stage_2, 'data/missing_incorrect_metadata_file_2.xlsx')

    def find_header(temp_df, excel_file):
        header_row = None
        for i, row in temp_df.iterrows():
            if row.str.contains('Document No').any():
                header_row = i
                break
        if header_row is not None:
            exportDocs = pd.read_excel(excel_file, header=header_row)
        else:
            raise ValueError("Header row with 'Document no' not found in the Export Document")
    
        return exportDocs


