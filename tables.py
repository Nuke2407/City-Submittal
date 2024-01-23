import customtkinter
from tkinter import filedialog, messagebox, ttk
import os
import shutil
import pandas as pd


def create_missing_incorrect_data_table_stage_1(self, excel_file_path):

    #Read Excel file into a dictionary of DataFrames - 'sheet_name=None' reads all sheets in the file.
    missing_incorrect_dfs = pd.read_excel(excel_file_path, sheet_name=None)

    #Initialize Dataframe.
    initial_sheet_name = next(iter(missing_incorrect_dfs))
    dataframe = missing_incorrect_dfs[initial_sheet_name]

    #Function to update the table with the selected sheet.
    def update_table(sheet_name):
        dataframe = missing_incorrect_dfs[sheet_name]

        #Simplified column names for Treeeview 
        simplified_columns = [col.replace('?', '').replace('(', '').replace(')', '').replace('/', '_').replace(' ', '_') for col in dataframe.columns]

        #clear the existing table data and columns.
        self.missing_incorrect_data_table.delete(*self.missing_incorrect_data_table.get_children())
        self.missing_incorrect_data_table['columns'] = simplified_columns

        #Create new columns based on the current dataframe
        for col, simplified_col_name in zip(dataframe.columns, simplified_columns):
            self.missing_incorrect_data_table.heading(simplified_col_name, text=col)
            self.missing_incorrect_data_table.column(simplified_col_name, anchor="center")

        #Insert new data.
        for _, row in dataframe.iterrows():
            values = [row[col] for col in dataframe.columns]
            self.missing_incorrect_data_table.insert("", "end", values=values)

    #Create a frame for table and description
    self.missing_incorrect_data_table_frame = customtkinter.CTkFrame(self.step_one_frame, corner_radius=0)
    self.missing_incorrect_data_table_frame.grid(row=5, column=0, padx=20, pady=10, sticky='new')
    self.missing_incorrect_data_table_frame.grid_rowconfigure(1, weight=1) 
    self.missing_incorrect_data_table_frame.grid_columnconfigure(0, weight=1)
    
    #Create a description for the table 
    self.missing_incorrect_data_table_description = customtkinter.CTkLabel(self.missing_incorrect_data_table_frame, text="The SQs displayed below have missing metadata. Please update in Aconex.",
                                                    font=customtkinter.CTkFont(size=15))
    self.missing_incorrect_data_table_description.grid(row=0, column=0, sticky='n')

    #Simplify column names and create the Treeview widget
    simplified_columns = [col.replace('?', '').replace('(', '').replace(')', '').replace('/', '_').replace(' ', '_') for col in dataframe.columns]
    self.missing_incorrect_data_table = ttk.Treeview(self.missing_incorrect_data_table_frame, columns=simplified_columns, show='headings')
    self.missing_incorrect_data_table.grid(row=1, column=0, padx=20, pady=10, sticky='nsew')

    #Create and place a vertical scrollbar
    scrollbar = ttk.Scrollbar(self.missing_incorrect_data_table_frame, orient="vertical", command=self.missing_incorrect_data_table.yview)
    scrollbar.grid(row=1, column=1, sticky='ns')
    self.missing_incorrect_data_table.configure(yscrollcommand=scrollbar.set)

    #Download Button for the Excel file.
    self.download_missing_metadata_file_button = customtkinter.CTkButton(self.missing_incorrect_data_table_frame, text="Download Missing Metadata Excel File", command=download_missing_incorrect_data_table)
    self.download_missing_metadata_file_button.grid(row=2, column=0, padx=20, pady=10)

    #Dropdown menu to select the sheet.
    sheet_names = list(missing_incorrect_dfs.keys())
    SQs_With_Missing_Data = customtkinter.StringVar(value=initial_sheet_name)
    sheet_selector = customtkinter.CTkOptionMenu(self.missing_incorrect_data_table_frame, values=sheet_names, variable=SQs_With_Missing_Data, command=update_table, corner_radius=0)
    sheet_selector.grid(row=2, column=0, sticky='wn', padx=20, pady=0)

    #Insert data into the table
    update_table(initial_sheet_name)

    if sheet_names:
        update_table(sheet_names[0])

def create_new_SQ_batch_table(self, excel_file_path):

    #Read Excel file into a dictionary of DataFrames - 'sheet_name=None' reads all sheets in the file.
    SQs_to_send_dfs = pd.read_excel(excel_file_path, sheet_name=None)

    #Initialize Dataframe.
    initial_sheet_name_2 = next(iter(SQs_to_send_dfs))
    dataframe = SQs_to_send_dfs[initial_sheet_name_2]

    #Function to update the table with the selected sheet.
    def update_table_SQs_to_send(sheet_name):
        dataframe = SQs_to_send_dfs[sheet_name]

        #Simplified column names for Treeeview 
        simplified_columns = [col.replace('?', '').replace('(', '').replace(')', '').replace('/', '_').replace(' ', '_') for col in dataframe.columns]

        #clear the existing table data and columns.
        self.SQs_to_send_data_table.delete(*self.SQs_to_send_data_table.get_children())
        self.SQs_to_send_data_table['columns'] = simplified_columns

        #Create new columns based on the current dataframe.
        for col, simplified_col_name in zip(dataframe.columns, simplified_columns):
            self.SQs_to_send_data_table.heading(simplified_col_name, text=col)
            self.SQs_to_send_data_table.column(simplified_col_name, anchor="center")

        #Insert new data.
        for _, row in dataframe.iterrows():
            values = [row[col] for col in dataframe.columns]
            self.SQs_to_send_data_table.insert("", "end", values=values)

    #Create a frame for table and description
    self.SQs_to_send_data_table_frame = customtkinter.CTkFrame(self.step_one_frame, corner_radius=0)
    self.SQs_to_send_data_table_frame.grid(row=6, column=0, padx=20, pady=10, sticky='new')
    self.SQs_to_send_data_table_frame.grid_rowconfigure(1, weight=1) 
    self.SQs_to_send_data_table_frame.grid_columnconfigure(0, weight=1)
    
    #Create a description for the table 
    self.SQs_to_send_data_table_description = customtkinter.CTkLabel(self.SQs_to_send_data_table_frame, text="The SQs displayed below will be sent to the City in this batch. Check to see if the metadata matches the PDF file.",
                                                    font=customtkinter.CTkFont(size=15))
    self.SQs_to_send_data_table_description.grid(row=0, column=0, sticky='n')

    #Simplify column names and create the Treeview widget
    simplified_columns = [col.replace('?', '').replace('(', '').replace(')', '').replace('/', '_').replace(' ', '_') for col in dataframe.columns]
    self.SQs_to_send_data_table = ttk.Treeview(self.SQs_to_send_data_table_frame, columns=simplified_columns, show='headings')
    self.SQs_to_send_data_table.grid(row=1, column=0, padx=20, pady=10, sticky='nsew')

    #Create and place a vertical scrollbar
    scrollbar = ttk.Scrollbar(self.SQs_to_send_data_table_frame, orient="vertical", command=self.SQs_to_send_data_table.yview)
    scrollbar.grid(row=1, column=1, sticky='ns')
    self.SQs_to_send_data_table.configure(yscrollcommand=scrollbar.set)

    #Download Button for the Excel file.
    self.download_missing_metadata_file_button = customtkinter.CTkButton(self.SQs_to_send_data_table_frame, text="Download SQ Checklist Excel File", command=download_SQs_to_send_file)
    self.download_missing_metadata_file_button.grid(row=2, column=0, padx=20, pady=10)

    #Dropdown menu to select the sheet.
    sheet_names_SQs_to_send = list(SQs_to_send_dfs.keys())
    SQs_to_send_Data = customtkinter.StringVar(value=initial_sheet_name_2)
    sheet_selector_SQs_to_send = customtkinter.CTkOptionMenu(self.SQs_to_send_data_table_frame, values=sheet_names_SQs_to_send, variable=SQs_to_send_Data, command=update_table_SQs_to_send, corner_radius=0)
    sheet_selector_SQs_to_send.grid(row=2, column=0, sticky='wn', padx=20, pady=0)

    #Insert data into the table
    update_table_SQs_to_send(initial_sheet_name_2)

    if sheet_names_SQs_to_send:
        update_table_SQs_to_send(sheet_names_SQs_to_send[0])

def create_missing_incorrect_data_table_stage_2(self, excel_file_path):

    #Read Excel file into a dictionary of DataFrame_stage_2s - 'sheet_name=None' reads all sheets in the file.
    missing_incorrect_dfs_stage_2 = pd.read_excel(excel_file_path, sheet_name=None)

    #Initialize Dataframe_stage_2.
    initial_sheet_name_stage_2 = next(iter(missing_incorrect_dfs_stage_2))
    dataframe_stage_2 = missing_incorrect_dfs_stage_2[initial_sheet_name_stage_2]

    #Function to update the table with the selected sheet.
    def update_table(sheet_name):
        dataframe_stage_2 = missing_incorrect_dfs_stage_2[sheet_name]

        #Simplified column names for Treeeview 
        simplified_columns = [col.replace('?', '').replace('(', '').replace(')', '').replace('/', '_').replace(' ', '_') for col in dataframe_stage_2.columns]

        #clear the existing table data and columns.
        self.missing_incorrect_data_table_stage_2.delete(*self.missing_incorrect_data_table_stage_2.get_children())
        self.missing_incorrect_data_table_stage_2['columns'] = simplified_columns

        #Create new columns based on the current dataframe_stage_2
        for col, simplified_col_name in zip(dataframe_stage_2.columns, simplified_columns):
            self.missing_incorrect_data_table_stage_2.heading(simplified_col_name, text=col)
            self.missing_incorrect_data_table_stage_2.column(simplified_col_name, anchor="center")

        #Insert new data.
        for _, row in dataframe_stage_2.iterrows():
            values = [row[col] for col in dataframe_stage_2.columns]
            self.missing_incorrect_data_table_stage_2.insert("", "end", values=values)

    #Create a frame for table and description
    self.missing_incorrect_data_table_stage_2_frame = customtkinter.CTkFrame(self.step_two_frame, corner_radius=0)
    self.missing_incorrect_data_table_stage_2_frame.grid(row=2, column=0, padx=20, pady=10, sticky='new')
    self.missing_incorrect_data_table_stage_2_frame.grid_rowconfigure(1, weight=1) 
    self.missing_incorrect_data_table_stage_2_frame.grid_columnconfigure(0, weight=1)

    #Create a description for the table 
    self.missing_incorrect_data_table_stage_2_description = customtkinter.CTkLabel(self.missing_incorrect_data_table_stage_2_frame, text="The SQs displayed below have missing metadata. Please update in Aconex.",
                                                    font=customtkinter.CTkFont(size=15))
    self.missing_incorrect_data_table_stage_2_description.grid(row=0, column=0, sticky='n')

    #Simplify column names and create the Treeview widget
    simplified_columns = [col.replace('?', '').replace('(', '').replace(')', '').replace('/', '_').replace(' ', '_') for col in dataframe_stage_2.columns]
    self.missing_incorrect_data_table_stage_2 = ttk.Treeview(self.missing_incorrect_data_table_stage_2_frame, columns=simplified_columns, show='headings')
    self.missing_incorrect_data_table_stage_2.grid(row=1, column=0, padx=20, pady=10, sticky='nsew')

    #Create and place a vertical scrollbar
    scrollbar = ttk.Scrollbar(self.missing_incorrect_data_table_stage_2_frame, orient="vertical", command=self.missing_incorrect_data_table_stage_2.yview)
    scrollbar.grid(row=1, column=1, sticky='ns')
    self.missing_incorrect_data_table_stage_2.configure(yscrollcommand=scrollbar.set)

    #Download Button for the Excel file.
    self.download_missing_metadata_file_button_stage_2 = customtkinter.CTkButton(self.missing_incorrect_data_table_stage_2_frame, text="Download Missing Metadata Excel File", command=download_missing_incorrect_data_table_stage_2)
    self.download_missing_metadata_file_button_stage_2.grid(row=2, column=0, padx=20, pady=10)

    #Dropdown menu to select the sheet.
    sheet_names_stage_2 = list(missing_incorrect_dfs_stage_2.keys())
    SQs_With_Missing_Data = customtkinter.StringVar(value=initial_sheet_name_stage_2)
    sheet_selector = customtkinter.CTkOptionMenu(self.missing_incorrect_data_table_stage_2_frame, values=sheet_names_stage_2, variable=SQs_With_Missing_Data, command=update_table, corner_radius=0)
    sheet_selector.grid(row=2, column=0, sticky='wn', padx=20, pady=0)

    #Insert data into the table
    update_table(initial_sheet_name_stage_2)

    if sheet_names_stage_2:
        update_table(sheet_names_stage_2[0])

def download_missing_incorrect_data_table():
    try:
        #Assuming the file is saved in the 'data' directory
        source_file = 'data/missing_incorrect_metadata_file.xlsx'
        if os.path.exists(source_file):
            # Ask the user where to save the file
            filetypes = [('Excel File', '*.xlsx'), ('All Files', '*.*')]
            dest_file = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=filetypes, title="Save File As")
            # If the user selects a location, copy the file
            if dest_file:
                shutil.copy(source_file, dest_file)             
            else:
                messagebox.showinfo("Download Cancelled", "Download operation was cancelled.")
        else:
            messagebox.showerror("Download Failed", "The source file does not exist.")
    
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

def download_SQs_to_send_file():
    try:
        #Assuming the file is saved in the 'data' directory
        source_file = 'data/new_city_sub_SQs.xlsx'
        if os.path.exists(source_file):
            # Ask the user where to save the file
            filetypes = [('Excel File', '*.xlsx'), ('All Files', '*.*')]
            dest_file = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=filetypes, title="Save File As")
            # If the user selects a location, copy the file
            if dest_file:
                shutil.copy(source_file, dest_file)             
            else:
                messagebox.showinfo("Download Cancelled", "Download operation was cancelled.")
        else:
            messagebox.showerror("Download Failed", "The source file does not exist.")
    
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

def download_missing_incorrect_data_table_stage_2():
    try:
        #Assuming the file is saved in the 'data' directory
        source_file = 'data/missing_incorrect_metadata_file_2.xlsx'
        if os.path.exists(source_file):
            # Ask the user where to save the file
            filetypes = [('Excel File', '*.xlsx'), ('All Files', '*.*')]
            dest_file = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=filetypes, title="Save File As")
            # If the user selects a location, copy the file
            if dest_file:
                shutil.copy(source_file, dest_file)             
            else:
                messagebox.showinfo("Download Cancelled", "Download operation was cancelled.")
        else:
            messagebox.showerror("Download Failed", "The source file does not exist.")
    
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")