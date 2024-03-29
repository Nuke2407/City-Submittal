import pandas as pd
import numpy as np
import os
from tkinter import messagebox
import re
from datetime import datetime, timedelta
import warnings

class Excel_Manipulation():
    def stage_1(self):
        self.import_manupulate_data()

        self.missing_metadata_file_1 = Excel_Manipulation.missing_metadata_check(self.SQ_metadata_closed_file)
        self.incorrect_metadata_file_stage_1 = Excel_Manipulation.incorrect_metadata_check(self.SQ_metadata_closed_file)

        #Create a downloadable version of the row that contains both the incorrect and missing metadata. 
        with pd.ExcelWriter('data/missing_incorrect_metadata_file.xlsx') as writer:
            self.missing_metadata_file_1.to_excel(writer, sheet_name='SQs With Missing Data', index=False)
            self.incorrect_metadata_file_stage_1.to_excel(writer, sheet_name='SQs With Incorrect Data', index=False)

        incorrect_missing_df_sheet_names = [(self.missing_metadata_file_1, 'SQs With Missing Data'), (self.incorrect_metadata_file_stage_1, 'SQs With Incorrect Data')]

        #Format the Excel file
        self.format_and_save_excel(incorrect_missing_df_sheet_names, 'data/missing_incorrect_metadata_file.xlsx')

    def import_manupulate_data(self):
        excel_file_1 = 'data/permanent/SQ_metadata_closed_template.xlsx'
        excel_file_2 = 'stage_one_documents/VLW-LOG-11000050-DC-0001_SQ_old.xls'
        excel_file_3 = 'stage_one_documents/ExportDocs_Step_1.xls'

        #Reads the excel file and stores the datat in rows. The 'header=' sets the header to the 12th excel column.
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

        #Filter the row to only include rows with 'N' in the 'If already sent to City?(Y/N)' column.
        self.new_SQs_for_city_submittal = self.SQ_metadata_closed_file.loc[
            (self.SQ_metadata_closed_file['If already sent to City?(Y/N)'] == 'N') & 
            (((self.SQ_metadata_closed_file['Design Directive'] == 'FCD – Field Change Directive') | 
            (self.SQ_metadata_closed_file['Design Directive'].isnull())) | 
            ((self.SQ_metadata_closed_file['Design Directive'] == 'RFI – No Design Change') & 
            ((self.SQ_metadata_closed_file['Category of Change'].isnull()) | 
            (self.SQ_metadata_closed_file['Class of Change'].isnull()))))]
            
        #Filter the row to only include rows with 'Y' in the 'If already sent to City?(Y/N)' column and rows with 'Y' in the 'Rev. updated?(Y/N/NA)' column.
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
        #Creates a new row where the Design Package, Incorporated To, and Specifications are split up and converted into lists for better comparison.
        incorrect_metadata_file = df_file.copy()
        pattern_combined = re.compile(r'(VLW-.*?)(?=, VLW-|$)')
        incorrect_metadata_file['Design Package Split Up'] = incorrect_metadata_file['Design Package'].apply(lambda x: pattern_combined.findall(x) if pd.notna(x) else '')
        incorrect_metadata_file['Specs Split Up'] = incorrect_metadata_file['Specifications'].apply(lambda x: pattern_combined.findall(x) if pd.notna(x) else '')
        incorrect_metadata_file['Incorporated Design Package Split Up'] = incorrect_metadata_file['Incorporated To'].apply(lambda x: pattern_combined.findall(x) if pd.notna(x) else '')
        incorrect_row = []

        #Iterating over the entire row to check the if the metadata has been fillout correctly.
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
                class_of_change_check = not row['Class of Change'] in ['No Action Required', 'Field Redline']
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
        incorrect_metadata_file = incorrect_metadata_file[['Document No', 'Title', 'Discipline', 'Category of Change', 'Design Directive', 'Class of Change', 
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
                'fg_color': '#1BF9F9',
                'border': 1
            })

            # Define a format for the data cells with borders
            cell_format = workbook.add_format({'border': 1})

            # Function to apply formatting to a worksheet
            def format_worksheet(worksheet, dataframe):
                # Write the column headers with the defined format
                for col_num, value in enumerate(dataframe.columns.values):
                    worksheet.write(0, col_num, value, header_format)

                    if value in ['Design Package', 'Incorporated To', 'Specifications', 'Title', 'Category of Change', 'Error Type']:
                        worksheet.set_column(col_num, col_num, 54)  
                    elif value in ['Reason for Change']:
                        worksheet.set_column(col_num, col_num, 41)                        
                    elif value in ['Document No', 'Design Directive', 'Class of Change']:
                        worksheet.set_column(col_num, col_num, 25)
                    elif value in ['Discipline']:
                        worksheet.set_column(col_num, col_num, 23)

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
                # Write the DataFrame data to XlsxWriter
                dataframe.to_excel(writer, sheet_name=sheet_name, index=False)
                # Apply formatting to each worksheet
                format_worksheet(writer.sheets[sheet_name], dataframe)
                (max_row, max_col) = dataframe.shape
                writer.sheets[sheet_name].autofilter(0, 0, max_row, max_col - 1)
                writer.sheets[sheet_name].freeze_panes(1, 0)

    def stage_2_part_1(self):
        excel_file_4 = 'stage_one_documents/ExportDocs_Step_2.xls' 
        temp_df_2 = pd.read_excel(excel_file_4, header=None)
        exportDocs_2 = Excel_Manipulation.find_header(temp_df_2, excel_file_4)

        excel_file_2 = 'stage_one_documents/ExportDocs_Step_1.xls' 
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

        incorrect_missing_df_sheet_names_stage_2 = [(missing_metadata_file_2, 'SQs With Missing Data 2'), (incorrect_metadata_file_stage_2, 'SQs With Incorrect Data 2')]

        #Format and create the Excel file
        self.format_and_save_excel(incorrect_missing_df_sheet_names_stage_2, 'data/missing_incorrect_metadata_file_2.xlsx')

    def stage_2_part_2(self):
        excel_file_5 = 'data/SQ_metadata_closed_file_2.xlsx' 
        temp_df_5 = pd.read_excel(excel_file_5, header=None)
        final_export_df_sheet_1 = Excel_Manipulation.find_header(temp_df_5, excel_file_5)
        
        excel_file_2 = 'stage_one_documents/VLW-LOG-11000050-DC-0001_SQ_old.xls'
        VLW_LOG_11000050_DC_0001_SQ_old = pd.read_excel(excel_file_2)

        self.new_batch_number = max(VLW_LOG_11000050_DC_0001_SQ_old['Batch #']) + 1 
        sqs_in_old_log = set(VLW_LOG_11000050_DC_0001_SQ_old['Document No'])
        final_export_df_FCD_only = final_export_df_sheet_1[final_export_df_sheet_1['Design Directive'].str.contains('FCD – Field Change Directive')].copy()
        final_export_df_FCD_only['Previous Batch #'] = ''
        final_export_df_FCD_only['Comment'] = ''
        final_export_df_FCD_only = final_export_df_FCD_only.rename(columns={'Incorporated To': 'Design Package(s) - Incorporated To'})

        sqs_with_issues = []
        def update_rows(row):
            if row['Class of Change'] in ['Field Redline', 'Not Incorporated in Design', 'Field Redline, Not Incorporated in Design', 'Not Incorporated in Design, Field Redline']:
                row['Design Package(s)'] = row['Design Package']
                row['Design Package(s) - Not Incorporated to'] = row['Design Package']

            elif row['Class of Change'] == 'Incorporated in Design':
                row['Design Package(s)'] = row['Design Package']
                row['Design Package(s) - Not Incorporated to'] = 'N/A'

            elif row['Class of Change'] in ['Partially Incorporated in Design', 'Field Redline, Incorporated in Design', 'Incorporated in Design, Field Redline']:
                row['Design Package(s)'] = str(row['Design Package']) + ', ' + str(row['Design Package(s) - Incorporated To'])
                row['Design Package(s) - Not Incorporated to'] = row['Design Package']
            else: 
                sqs_with_issues.append(row)

            if row['Document No'] in sqs_in_old_log:
                row['Batch #'] = VLW_LOG_11000050_DC_0001_SQ_old.loc[VLW_LOG_11000050_DC_0001_SQ_old['Document No'] == row['Document No'], 'Batch #'].iloc[0]
            else: 
                row['Batch #'] = self.new_batch_number

            return row
        
        final_export_df_FCD_only = final_export_df_FCD_only.apply(update_rows, axis=1)
        city_submittal_sheet_1_formated = final_export_df_FCD_only[['Document No', 'Revision', 'Title', 'Discipline', 'Design Package(s)', 'Design Package(s) - Not Incorporated to', 'Design Package(s) - Incorporated To', 'Specifications', 'Category of Change', 'Revision Date',	'Design Directive', 'Status', 'Class of Change', 'Batch #',	'Previous Batch #', 'Comment']]

        #Add SQs that are now open from the previous log excluding the superseded 'Cancelled' ones. 
        new_log_identifiers = city_submittal_sheet_1_formated['Document No']
        sqs_to_add_back = VLW_LOG_11000050_DC_0001_SQ_old[~VLW_LOG_11000050_DC_0001_SQ_old['Document No'].isin(new_log_identifiers)]
        closed_sqs_to_add_back = sqs_to_add_back[sqs_to_add_back['Status'].str.contains('Closed')].copy()
        city_submittal_sheet_1_formated_sqs_added = pd.concat([city_submittal_sheet_1_formated, closed_sqs_to_add_back])

        excel_file_7 = 'data/new_city_sub_SQs.xlsx'
        if os.path.exists(excel_file_7):
            sqs_reved_up = pd.read_excel(excel_file_7, sheet_name='Reved up SQs', header=0)
            if not sqs_reved_up.empty:
                for index, row in sqs_reved_up.iterrows():
                    city_submittal_sheet_1_formated_sqs_added.loc[city_submittal_sheet_1_formated_sqs_added['Document No'] == row['Document No'], 'Previous Batch #'] = city_submittal_sheet_1_formated_sqs_added.loc[city_submittal_sheet_1_formated_sqs_added['Document No'] == row['Document No'], 'Batch #']
                    city_submittal_sheet_1_formated_sqs_added.loc[city_submittal_sheet_1_formated_sqs_added['Document No'] == row['Document No'], 'Batch #'] = self.new_batch_number
        else:
            messagebox.showerror("Error", "For some reason 'new_city_sub_SQs.xlsx' is missing. Try restarting the entire City Submittal process.")

        #Delete on hold SQs if not in the previous log or add the version that was sent in the previous city sub. Only executes function if a sq_removal_excel file was uploaded. 
        excel_file_6 = 'stage_one_documents/sq_removal_excel.xlsx' 
        if os.path.exists(excel_file_6):
            sqs_on_hold_df = pd.read_excel(excel_file_6, sheet_name='SQs on hold', header=0)
            sqs_to_supersed_df = pd.read_excel(excel_file_6, sheet_name='SQs to supersed', header=0)
            if not sqs_on_hold_df.empty:
                sqs_to_delete = sqs_on_hold_df['Document Number'].to_list()
                city_submittal_sheet_1_formated_on_hold_sqs_removed = city_submittal_sheet_1_formated_sqs_added[~city_submittal_sheet_1_formated_sqs_added['Document No'].isin(sqs_to_delete)]
                sqs_to_add_back_2 = VLW_LOG_11000050_DC_0001_SQ_old[VLW_LOG_11000050_DC_0001_SQ_old['Document No'].isin(sqs_to_delete)]
                city_submittal_sheet_1_formated_sqs_added = pd.concat([city_submittal_sheet_1_formated_on_hold_sqs_removed, sqs_to_add_back_2])
            if not sqs_to_supersed_df.empty:
                for index, row in sqs_to_supersed_df.iterrows(): 
                    superseding_sq_number = row['Document Number of superseding SQ'][-4:]
                    superseding_sq_revision_number = str(row['Revision Number of superseding SQ'])[-1:]
                    city_submittal_sheet_1_formated_sqs_added.loc[city_submittal_sheet_1_formated_sqs_added['Document No'] == row['Document Number of SQ being superseded'], 'Comment'] = f"This SQ is obsolete and superseded by SQ-{superseding_sq_number} REV 0{superseding_sq_revision_number}. It will be permanently deleted starting from City Submittal #{self.new_batch_number + 1}"
                    city_submittal_sheet_1_formated_sqs_added.loc[city_submittal_sheet_1_formated_sqs_added['Document No'] == row['Document Number of SQ being superseded'], 'Status'] = 'Cancelled'
                    city_submittal_sheet_1_formated_sqs_added.loc[city_submittal_sheet_1_formated_sqs_added['Document No'] == row['Document Number of SQ being superseded'], 'Previous Batch #'] = city_submittal_sheet_1_formated_sqs_added.loc[city_submittal_sheet_1_formated_sqs_added['Document No'] == row['Document Number of SQ being superseded'], 'Batch #']
                    city_submittal_sheet_1_formated_sqs_added.loc[city_submittal_sheet_1_formated_sqs_added['Document No'] == row['Document Number of SQ being superseded'], 'Batch #'] = self.new_batch_number
	
        city_submittal_sheet_1 = city_submittal_sheet_1_formated_sqs_added.sort_values(by='Document No')

        #SHEET 4
        #=====================================================================================================================

        city_submittal_sheet_4_packages_organized = city_submittal_sheet_1_formated_sqs_added.copy()
        pattern_combined = re.compile(r'(VLW-.*?)(?=, VLW-|$)')
        # Use a lambda function with findall to get all matches in a list directly
        city_submittal_sheet_4_packages_organized['Design Package Split Up'] = city_submittal_sheet_4_packages_organized['Design Package(s)'].apply(lambda x: pattern_combined.findall(str(x)))
        city_submittal_sheet_4_packages_organized['DP Not Incorporated to Split Up'] = city_submittal_sheet_4_packages_organized['Design Package(s) - Not Incorporated to'].apply(lambda x: pattern_combined.findall(str(x)))
        city_submittal_sheet_4_packages_organized['DP Incorporated To Split Up'] = city_submittal_sheet_4_packages_organized['Design Package(s) - Incorporated To'].apply(lambda x: pattern_combined.findall(str(x)))
        sheet_4_sort_df = city_submittal_sheet_4_packages_organized[['Design Package Split Up', 'DP Not Incorporated to Split Up', 'DP Incorporated To Split Up', 'Category of Change', 'Class of Change']]

        city_submittal_sheet_4 = pd.DataFrame(columns=['Design package number', 'Design Package Title', 'Field Redline', 'Incorporated in Design', 'Not Incorporated in Design', 'Major Change (significant change - requires City review)', 
                    'Minor Change (not a significant change to design intent)', 'Urgent/Unforeseen Change (in-progress construct. activities)'])
        
        pattern = r"^(- |– |-)"

        for index, row in sheet_4_sort_df.iterrows():
            if isinstance(row['DP Not Incorporated to Split Up'], list) and row['DP Not Incorporated to Split Up']:
                for doc_num_title in row['DP Not Incorporated to Split Up']:
                    if doc_num_title.startswith("VLW-SPC") or doc_num_title.startswith("VLW-DNR"):
                        continue 
                    doc_num, doc_title = doc_num_title.split(' ', maxsplit=1)
                    if doc_num not in city_submittal_sheet_4['Design package number'].tolist():
                        doc_title_short = re.sub(pattern, '', doc_title)
                        new_row = {
                        'Design package number': doc_num, 
                        'Design Package Title': doc_title_short,
                        'Field Redline': 0, 
                        'Incorporated in Design': 0, 
                        'Not Incorporated in Design': 0,
                        'Major Change (significant change - requires City review)': 0, 
                        'Minor Change (not a significant change to design intent)': 0, 
                        'Urgent/Unforeseen Change (in-progress construct. activities)': 0}
                        new_row_df = pd.DataFrame([new_row])
                        city_submittal_sheet_4 = pd.concat([city_submittal_sheet_4, new_row_df], ignore_index=True)

                    match_index = city_submittal_sheet_4.index[city_submittal_sheet_4['Design package number'] == doc_num].tolist()
                    if row['Class of Change'] in ['Field Redline', 'Field Redline, Incorporated in Design', 'Incorporated in Design, Field Redline']:
                        if doc_num == 'VLW-PKG-00000036-FC-0001':
                            pd.DataFrame(row).to_excel('data/BUGSSSS.xlsx')
                            print('BREAK')
                            print(new_row)
                        city_submittal_sheet_4.at[match_index[0], 'Field Redline'] += 1
                    if row['Class of Change'] in ['Not Incorporated in Design', 'Partially Incorporated in Design', 'Field Redline, Not Incorporated in Design', 'Not Incorporated in Design, Field Redline']:
                        city_submittal_sheet_4.at[match_index[0], 'Not Incorporated in Design'] += 1
                        if row['Category of Change'] == 'Minor Change (not a significant change to design intent)':
                            city_submittal_sheet_4.at[match_index[0], 'Minor Change (not a significant change to design intent)'] += 1
                        if row['Category of Change'] == 'Major Change (significant change - requires City review)':
                            city_submittal_sheet_4.at[match_index[0], 'Major Change (significant change - requires City review)'] += 1
                        if row['Category of Change'] == 'Urgent/Unforeseen Change (in-progress construct. activities)':
                            city_submittal_sheet_4.at[match_index[0], 'Urgent/Unforeseen Change (in-progress construct. activities)'] += 1

            if isinstance(row['DP Incorporated To Split Up'], list) and row['DP Incorporated To Split Up']:
                for doc_num_title in row['DP Incorporated To Split Up']:
                    if doc_num_title.startswith("VLW-SPC") or doc_num_title.startswith("VLW-DNR"):
                        continue 
                    doc_num, doc_title = doc_num_title.split(' ', maxsplit=1)
                    if doc_num not in city_submittal_sheet_4['Design package number'].tolist():
                        doc_title_short = re.sub(pattern, '', doc_title)
                        new_row = {
                        'Design package number': doc_num, 
                        'Design Package Title': doc_title_short,
                        'Field Redline': 0, 
                        'Incorporated in Design': 0, 
                        'Not Incorporated in Design': 0,
                        'Major Change (significant change - requires City review)': 0, 
                        'Minor Change (not a significant change to design intent)': 0, 
                        'Urgent/Unforeseen Change (in-progress construct. activities)': 0}
                        new_row_df = pd.DataFrame([new_row])
                        city_submittal_sheet_4 = pd.concat([city_submittal_sheet_4, new_row_df], ignore_index=True)

                    match_index = city_submittal_sheet_4.index[city_submittal_sheet_4['Design package number'] == doc_num].tolist()
                    if row['Class of Change'] in ['Partially Incorporated in Design', 'Incorporated in Design', 'Field Redline, Incorporated in Design', 'Incorporated in Design, Field Redline']:
                        city_submittal_sheet_4.at[match_index[0], 'Incorporated in Design'] += 1

        city_submittal_sheet_4 = city_submittal_sheet_4.sort_values(by= 'Design package number')

        #Sheet #2
        #=========================================================================================================================================================================================================
        
        #Table #1
        city_submittal_sheet_2_setup_table_1 = pd.DataFrame(columns= ['Discipline', 'SQ Count'])
        for index, row in city_submittal_sheet_1_formated_sqs_added.iterrows():
            if row['Status'] == 'Closed':
                if row['Discipline'] not in city_submittal_sheet_2_setup_table_1['Discipline'].tolist():
                    new_row = {
                            'Discipline': row['Discipline'], 
                            'SQ Count': 0}
                    new_row_df = pd.DataFrame([new_row])
                    city_submittal_sheet_2_setup_table_1 = pd.concat([city_submittal_sheet_2_setup_table_1, new_row_df], ignore_index=True)

                match_index = city_submittal_sheet_2_setup_table_1.index[city_submittal_sheet_2_setup_table_1['Discipline'] == row['Discipline']].tolist()
                if row['Discipline'] in city_submittal_sheet_2_setup_table_1['Discipline'].tolist():
                    city_submittal_sheet_2_setup_table_1.at[match_index[0], 'SQ Count'] += 1
        
        total_df = pd.DataFrame([{'Discipline': 'Total', 'SQ Count': city_submittal_sheet_2_setup_table_1['SQ Count'].sum()}])
        city_submittal_sheet_2_setup_table_1 = pd.concat([city_submittal_sheet_2_setup_table_1, total_df], ignore_index=True)
        
        #Table #2
        city_submittal_sheet_2_setup_table_2= pd.DataFrame({
            'Category of Change':[
                'Minor Change (not a significant change to design intent)', 
                'Urgent/Unforeseen Change (in-progress construct. activities)', 
                'Major Change (significant change - requires City review)'], 
            'SQ Count':[0,0,0], 
            'Incorporated in Design':[0,0,0], 
            'Not Incorporated in Design':[0,0,0], 
            'Partially Incorporated in Design':[0,0,0], 
            'Field Redline':[0,0,0]})
        
        def sheet_2_class_of_change_count(row, num):
            if row['Class of Change'] == 'Incorporated in Design':
                city_submittal_sheet_2_setup_table_2.loc[num, 'Incorporated in Design'] += 1
            if row['Class of Change'] in ['Not Incorporated in Design', 'Field Redline, Not Incorporated in Design']:
                city_submittal_sheet_2_setup_table_2.loc[num, 'Not Incorporated in Design'] += 1
            if row['Class of Change'] == 'Partially Incorporated in Design':
                city_submittal_sheet_2_setup_table_2.loc[num, 'Partially Incorporated in Design'] += 1
            if row['Class of Change'] in ['Field Redline', 'Field Redline, Incorporated in Design']:  
                city_submittal_sheet_2_setup_table_2.loc[num, 'Field Redline'] += 1
        
        for index, row in city_submittal_sheet_1_formated_sqs_added.iterrows():
            if row['Status'] == 'Closed':
                if row['Category of Change'] == 'Minor Change (not a significant change to design intent)':
                    city_submittal_sheet_2_setup_table_2.loc[0, 'SQ Count'] += 1 
                    sheet_1_row_num = 0
                    sheet_2_class_of_change_count(row, sheet_1_row_num)

                elif row['Category of Change'] == 'Urgent/Unforeseen Change (in-progress construct. activities)':
                    city_submittal_sheet_2_setup_table_2.loc[1, 'SQ Count'] += 1 
                    sheet_1_row_num = 1
                    sheet_2_class_of_change_count(row, sheet_1_row_num)

                elif row['Category of Change'] == 'Major Change (significant change - requires City review)':
                    city_submittal_sheet_2_setup_table_2.loc[2, 'SQ Count'] += 1 
                    sheet_1_row_num = 2
                    sheet_2_class_of_change_count(row, sheet_1_row_num)


        
        sq_count_sum = city_submittal_sheet_2_setup_table_2['SQ Count'].sum()
        incorporated_in_desing_sum = city_submittal_sheet_2_setup_table_2['Incorporated in Design'].sum()
        not_incorporated_in_design_sum = city_submittal_sheet_2_setup_table_2['Not Incorporated in Design'].sum()
        partially_incorporated_in_design_sum = city_submittal_sheet_2_setup_table_2['Partially Incorporated in Design'].sum()
        field_redline_sum = city_submittal_sheet_2_setup_table_2['Field Redline'].sum()

        total_df = pd.DataFrame([{'Category of Change': 'Total', 'SQ Count': sq_count_sum, 'Incorporated in Design': incorporated_in_desing_sum, 'Not Incorporated in Design':not_incorporated_in_design_sum, 'Partially Incorporated in Design': partially_incorporated_in_design_sum, 'Field Redline': field_redline_sum}])
        city_submittal_sheet_2_setup_table_2 = pd.concat([city_submittal_sheet_2_setup_table_2, total_df], ignore_index=True)

        #Sheet #3
        #=========================================================================================================================================================================================================
        
        #Ignore the 
        warnings.filterwarnings(action='ignore', category=FutureWarning, 
                        message=".*DataFrame concatenation with empty or all-NA entries is deprecated.*")

        city_submittal_sheet_3 = pd.DataFrame(columns=['Design package number', 'Document No', 'Revision', 'Title', 'Discipline', 'Specifications', 'Category of Change',  'Revision Date', 'Design Directive', 'Status', 'Class of Change'])

        for index, row in city_submittal_sheet_4_packages_organized.iterrows():
            if row['Status'] != 'Closed':
                continue
            if isinstance(row['Design Package Split Up'], list) and row['Design Package Split Up']:
                for doc_num_title in row['Design Package Split Up']:
                    if doc_num_title.startswith("VLW-SPC") or doc_num_title.startswith("VLW-DNR"):
                        continue 
                    doc_num, doc_title = doc_num_title.split(' ', maxsplit=1)
                    new_row = {
                        'Design package number': doc_num,
                        'Document No': row['Document No'],
                        'Revision': row['Revision'],
                        'Title': row['Title'],
                        'Discipline': row['Discipline'],
                        'Specifications': row['Specifications'],
                        'Category of Change': row['Category of Change'],
                        'Revision Date': row['Revision Date'],
                        'Design Directive': row['Design Directive'],
                        'Status': row['Status'],
                        'Class of Change': row['Class of Change']}
                    
                    new_row_df = pd.DataFrame([new_row], columns=city_submittal_sheet_3.columns)
                    city_submittal_sheet_3 = pd.concat([city_submittal_sheet_3, new_row_df], ignore_index=True)

        city_submittal_sheet_3 = city_submittal_sheet_3.sort_values(by= 'Design package number')

        friday_str = str(self.upcoming_friday())
        friday_date = friday_str.split(' ', maxsplit=1)[0]  # Take the first part of the split
        friday_date = 'Summary ' + friday_date  # Concatenate with 'Summary'
        
        dataframes_with_sheet_1_info = [(city_submittal_sheet_1, f'Batch#{self.new_batch_number}')]
        dataframes_with_sheet_2_info = [(city_submittal_sheet_2_setup_table_1, friday_date), (city_submittal_sheet_2_setup_table_2, friday_date)]
        dataframes_with_sheet_4_info = [(city_submittal_sheet_4, 'Summary-Design Package Tracker')]
        dataframes_with_sheet_3_info = [(city_submittal_sheet_3, 'Design Package Tracker')]
        self.write_to_excel('data/City Submittal.xlsx', dataframes_with_sheet_1_info, dataframes_with_sheet_2_info, dataframes_with_sheet_4_info, dataframes_with_sheet_3_info)

    def write_to_excel(self, file_path, dataframes_for_sheet_1, dataframes_for_sheet_2, dataframes_for_sheet_4, dataframes_for_sheet_3):
        #Initialize the ExcelWriter once
        with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
            #Call the sheet formatting and writing functions
            self.format_sheet_1_final_export(writer, dataframes_for_sheet_1)
            self.format_sheet_2_final_export(writer, dataframes_for_sheet_2)
            self.format_sheet_3_final_export(writer, dataframes_for_sheet_3)
            self.format_sheet_4_final_export(writer, dataframes_for_sheet_4)

    def format_sheet_1_final_export(self, writer, dataframes_with_sheet_names):
            
            # Access the workbook and configure it
            workbook = writer.book
            workbook.nan_inf_to_errors = True

            # Define a format for the header
            header_format = workbook.add_format({
                'font_size': 10,
                'bold': True,
                'text_wrap': True,
                'valign': 'vcenter',
                'align': 'center',
                'font_color': '#00008B',
                'border': 1})
            # Define a format for the data cells with borders
            cell_format = workbook.add_format({
                'font_size': 10,
                'text_wrap': True,
                'valign': 'vcenter',
                'align': 'center',
                'border': 1,})
            date_format = workbook.add_format({
                'font_size': 10,
                'text_wrap': True,
                'valign': 'vcenter',
                'align': 'center',
                'border': 1,
                'num_format': 'mm/dd/yyyy'})
            yellow_format = workbook.add_format({
                'font_size': 10, 
                'bg_color': '#FFFF00', 
                'text_wrap': True, 
                'valign': 'vcenter', 
                'align': 'center', 
                'border': 1})  
            yellow_format_date = workbook.add_format({
                'font_size': 10, 
                'bg_color': '#FFFF00', 
                'text_wrap': True, 
                'valign': 'vcenter', 
                'align': 'center', 
                'border': 1, 
                'num_format': 'mm/dd/yyyy'})
            red_format = workbook.add_format({
                'font_size': 10, 
                'bg_color': '#FF0000', 
                'text_wrap': True, 
                'valign': 'vcenter', 
                'align': 'center', 
                'border': 1}) 
            red_format_date = workbook.add_format({
                'font_size': 10, 
                'bg_color': '#FF0000', 
                'text_wrap': True, 
                'valign': 'vcenter', 
                'align': 'center', 
                'border': 1, 
                'num_format': 'mm/dd/yyyy'})    

            # Function to apply formatting to a worksheet
            def format_worksheet(worksheet, dataframe):
                # Write the column headers with the defined format
                for col_num, value in enumerate(dataframe.columns.values):
                    worksheet.write(0, col_num, value, header_format)
                    worksheet.set_zoom(70)

                    # Set specific column widths
                    if value in ['Design Package(s)', 'Design Package(s) - Not Incorporated to', 'Design Package(s) - Incorporated To', 'Specifications']:
                        worksheet.set_column(col_num, col_num, 54)  
                    elif value in ['Document No', 'Comment', 'Title']:
                        worksheet.set_column(col_num, col_num, 30)
                    elif value in ['Discipline', 'Category of Change', 'Design Directive', 'Class of Change']:
                        worksheet.set_column(col_num, col_num, 21)
                    elif value in ['Revision', 'Revision Date', 'Status', 'Batch', 'Previous Batch #']:
                        worksheet.set_column(col_num, col_num, 11)

                # Set row height to 69 pixels (approximately 51 points)
                worksheet.set_default_row(80)

                status_col_index = dataframe.columns.get_loc("Status") if "Status" in dataframe.columns else None
                previous_batch_col_index = dataframe.columns.get_loc("Previous Batch #") if "Previous Batch #" in dataframe.columns else None

                # Apply the cell format to each cell in the data
                for row_num in range(len(dataframe)):
                    status_value = dataframe.iloc[row_num, status_col_index] if status_col_index is not None else ""
                    previous_batch_value = dataframe.iloc[row_num, previous_batch_col_index] if previous_batch_col_index is not None else None

                    # Loop through each column for the current row
                    for col_num, column_name in enumerate(dataframe.columns):
                        cell_value = dataframe.iloc[row_num, col_num]
                        if pd.isna(cell_value):
                            cell_value = None  # Handle NaN and NaT

                        # Determine the base format
                        base_format = cell_format
                        if column_name == 'Revision Date':
                            base_format = date_format

                        # Determine conditional formatting based on 'Status' and 'Previous Batch #'
                        if status_value == 'Cancelled':
                            format_to_use = red_format_date if column_name == 'Revision Date' else red_format
                        elif status_value == 'Closed' and isinstance(previous_batch_value, int):
                            format_to_use = yellow_format_date if column_name == 'Revision Date' else yellow_format
                        else:
                            format_to_use = base_format

                        # Write the cell with the determined format
                        worksheet.write(row_num + 1, col_num, cell_value, format_to_use)

            # Write each DataFrame to its respective sheet and apply formatti
            for dataframe, sheet_name in dataframes_with_sheet_names:
                # Write the DataFrame data to XlsxWriter
                dataframe.to_excel(writer, sheet_name=sheet_name, index=False)
                # Apply formatting to each worksheet
                format_worksheet(writer.sheets[sheet_name], dataframe)

                if sheet_name == f'Batch#{self.new_batch_number}':
                    (max_row, max_col) = dataframe.shape
                    writer.sheets[sheet_name].autofilter(0, 0, max_row, max_col - 1)
                    writer.sheets[sheet_name].freeze_panes(1, 0)

    def format_sheet_3_final_export(self, writer, dataframes_with_sheet_names):
            
            # Access the workbook and configure it
            workbook = writer.book
            workbook.nan_inf_to_errors = True

            # Define a format for the header
            header_format = workbook.add_format({
                'font_size': 10,
                'bold': True,
                'text_wrap': True,
                'valign': 'vcenter',
                'align': 'center',
                'font_color': '#00008B',
                'border': 1})
            # Define a format for the data cells with borders
            cell_format = workbook.add_format({
                'font_size': 10,
                'text_wrap': True,
                'valign': 'vcenter',
                'align': 'center',
                'border': 1,})
            date_format = workbook.add_format({
                'font_size': 10,
                'text_wrap': True,
                'valign': 'vcenter',
                'align': 'center',
                'border': 1,
                'num_format': 'mm/dd/yyyy'})

            # Function to apply formatting to a worksheet
            def format_worksheet(worksheet, dataframe):
                # Write the column headers with the defined format
                for col_num, value in enumerate(dataframe.columns.values):
                    worksheet.write(0, col_num, value, header_format)
                    worksheet.set_zoom(70)

                    # Set specific column widths
                    if value in ['Title']:
                        worksheet.set_column(col_num, col_num, 82)                     
                    if value in ['Specifications', 'Class of Change', 'Category of Change']:
                        worksheet.set_column(col_num, col_num, 54)  
                    if value in ['Class of Change']:
                        worksheet.set_column(col_num, col_num, 36)                         
                    elif value in ['Design package number', 'Document No', 'Design Directive']:
                        worksheet.set_column(col_num, col_num, 30)
                    elif value in ['Discipline']:
                        worksheet.set_column(col_num, col_num, 21)
                    elif value in ['Revision Date']:
                        worksheet.set_column(col_num, col_num, 17)
                    elif value in ['Revision', 'Status']:
                        worksheet.set_column(col_num, col_num, 11)

                # Set row height to 69 pixels (approximately 51 points)
                worksheet.set_default_row(19)

                # Apply the cell format to each cell in the data
                for row_num in range(len(dataframe)):

                    # Loop through each column for the current row
                    for col_num, column_name in enumerate(dataframe.columns):
                        cell_value = dataframe.iloc[row_num, col_num]
                        if pd.isna(cell_value):
                            cell_value = None  # Handle NaN and NaT

                        # Determine the base format
                        base_format = cell_format
                        if column_name == 'Revision Date':
                            base_format = date_format

                        # Write the cell with the determined format
                        worksheet.write(row_num + 1, col_num, cell_value, base_format)

            # Write each DataFrame to its respective sheet and apply formatti
            for dataframe, sheet_name in dataframes_with_sheet_names:
                # Write the DataFrame data to XlsxWriter
                dataframe.to_excel(writer, sheet_name=sheet_name, index=False)
                # Apply formatting to each worksheet
                format_worksheet(writer.sheets[sheet_name], dataframe)
                (max_row, max_col) = dataframe.shape
                writer.sheets[sheet_name].autofilter(0, 0, max_row, max_col - 1)
                writer.sheets[sheet_name].freeze_panes(1, 0)

    def format_sheet_2_final_export(self, writer, dataframes_with_sheet_names):
        df1, sheet_name1 = dataframes_with_sheet_names[0]
        df2, sheet_name2 = dataframes_with_sheet_names[1]

        # Access the workbook
        workbook = writer.book
        worksheet = workbook.add_worksheet(name=sheet_name1)  # Assuming both DataFrames go to the same named sheet

        # Define a base format for headers, similar to your existing format
        header_format = workbook.add_format({
            'font_size': 10,
            'bold': True,
            'text_wrap': True,
            'valign': 'vcenter',
            'align': 'center',
            'border': 1})
        merged_format = workbook.add_format({
            'font_size': 10,
            'bold': True,
            'text_wrap': True,
            'valign': 'vcenter',
            'align': 'center',
            'border': 1})
        cell_format = workbook.add_format({
            'font_size': 10,
            'bold': True,
            'text_wrap': True,
            'valign': 'vcenter',
            'align': 'center',
            'border': 1})

        # Function to write a DataFrame to a specified starting row
        def write_dataframe(worksheet, dataframe, start_row):
            # Write the headers
            for col_num, value in enumerate(dataframe.columns.values):
                worksheet.write(start_row, col_num, value, header_format)

                if value in ['Discipline']:
                    worksheet.set_column(col_num, col_num, 52)  
                elif value in ['SQ Count']:
                    worksheet.set_column(col_num, col_num, 10)
                elif value in ['Field Redline', 'Incorporated in Design', 'Not Incorporated in Design', 'Partially Incorporated in Design']:
                    worksheet.set_column(col_num, col_num, 27)
            
            # Write the data
            for row_num, row in enumerate(dataframe.values, start=start_row + 1):
                for col_num, cell in enumerate(row):
                    worksheet.write(row_num, col_num, cell, cell_format)  

        # Write the first DataFrame
        write_dataframe(worksheet, df1, 0)  # Starting from row 0

        # Calculate the starting row for the second DataFrame (df1's rows + 2 for spacing)
        start_row_df2 = len(df1) + 3  
        row_merge_num = len(df1) + 3
        row_merge = f"C{row_merge_num}:F{row_merge_num}"
        worksheet.merge_range(row_merge, 'Class of Change', merged_format)

        # Write the second DataFrame
        write_dataframe(worksheet, df2, start_row_df2)

    def format_sheet_4_final_export(self, writer, dataframes_with_sheet_names):
            # Access the workbook and configure it
            workbook = writer.book
            workbook.nan_inf_to_errors = True

            # Define a format for the header
            header_format = workbook.add_format({
                'font_size': 10,
                'bold': True,
                'text_wrap': True,
                'valign': 'vcenter',
                'align': 'center',
                'border': 1})
            # Define a format for the merged cell
            merged_format_blue = workbook.add_format({
                'font_size': 10,
                'bold': True,
                'align': 'center',
                'valign': 'vcenter',
                'bg_color': '#DDEBF7',
                'border': 1})
            merged_format_green = workbook.add_format({
                'font_size': 10,
                'bold': True,
                'align': 'center',
                'valign': 'vcenter',
                'bg_color': '#E2EFDA',
                'border': 1})
            # Define a format for the data cells with borders
            cell_format = workbook.add_format({
                'font_size': 10,
                'text_wrap': True,
                'valign': 'vcenter',
                'align': 'center',
                'border': 1,})
            start_row = 1
            #Function to apply formatting to a worksheet
            def format_worksheet(worksheet, dataframe):
                # Merge cells in the first row for the 'Good' and 'Bad' categories
                worksheet.merge_range('C1:E1', 'Class of Change', merged_format_blue)
                worksheet.merge_range('F1:H1', 'Category of Change for Not Incorporated SQs', merged_format_green)

                #Write the column headers with the defined format
                for col_num, value in enumerate(dataframe.columns.values):
                    worksheet.write(start_row, col_num, value, header_format)
                    worksheet.set_zoom(100)
                    # Set specific column widths
                    if value in ['Design package number']:
                        worksheet.set_column(col_num, col_num, 26)  
                    elif value in ['Design Package Title']:
                        worksheet.set_column(col_num, col_num, 33)
                    elif value in ['Field Redline', 'Incorporated in Design', 'Not Incorporated in Design', 'Major Change (significant change - requires City review)', 'Minor Change (not a significant change to design intent)', 'Urgent/Unforeseen Change (in-progress construct. activities)']:
                        worksheet.set_column(col_num, col_num, 24)
                
                # Set the height of the header row to 27
                worksheet.set_row(start_row, 27)
                #Set row height to 69 pixels (approximately 12.75 points)
                worksheet.set_default_row(12.75)

                # Apply the cell format to each cell in the data
                for row_num, row in enumerate(dataframe.values, start=start_row + 1):
                    for col_num, cell_value in enumerate(row):
                        # Check for NaT values and replace with None
                        cell_value = None if pd.isna(cell_value) else cell_value
                        worksheet.write(row_num, col_num, cell_value, cell_format)

            # Write each DataFrame to its respective sheet and apply formatti
            for dataframe, sheet_name in dataframes_with_sheet_names:
                # Write the DataFrame data to XlsxWriter
                worksheet = workbook.add_worksheet(sheet_name)
                # Apply formatting to each worksheet
                format_worksheet(worksheet, dataframe)

                if sheet_name == 'Summary-Design Package Tracker':
                    (max_row, max_col) = dataframe.shape
                    writer.sheets[sheet_name].autofilter(1, 0, max_row + start_row, max_col - 1)
                    writer.sheets[sheet_name].freeze_panes(2, 0)

    def find_header(temp_df, excel_file):
        header_row = None
        for i, row in temp_df.iterrows():
            if (row == 'Document No').any():
                header_row = i
                break
        if header_row is not None:
            exportDocs = pd.read_excel(excel_file, header=header_row)
        else:
            raise ValueError("Header row with 'Document no' not found in the Export Document")
    
        return exportDocs

    def upcoming_friday(self):
        # Today's date
        today = datetime.today()

        # Weekday of today, where Monday is 0 and Sunday is 6
        today_weekday = today.weekday()

        # Days until next Friday (4 represents Friday)
        days_until_next_friday = (4 - today_weekday) % 7

        # Next Friday's date
        next_friday = today + timedelta(days=days_until_next_friday)

        # Output next Friday's date
        return next_friday
