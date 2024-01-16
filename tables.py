import customtkinter
from tkinter import filedialog, messagebox, ttk
import os
import shutil
import pandas as pd


def create_missing_incorrect_data_table(app, excel_file_path):

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
        app.missing_data_table.delete(*app.missing_data_table.get_children())
        app.missing_data_table['columns'] = simplified_columns

        #Create new columns based on the current dataframe.
        for col, simplified_col_name in zip(dataframe.columns, simplified_columns):
            app.missing_data_table.heading(simplified_col_name, text=col)
            app.missing_data_table.column(simplified_col_name, anchor="center")

        #Insert new data.
        for _, row in dataframe.iterrows():
            values = [row[col] for col in dataframe.columns]
            app.missing_data_table.insert("", "end", values=values)

    #Create a frame for table and description
    app.missing_data_table_frame = customtkinter.CTkFrame(app.step_one_frame, corner_radius=0)
    app.missing_data_table_frame.grid(row=5, column=0, padx=20, pady=10, sticky='new')
    app.missing_data_table_frame.grid_rowconfigure(1, weight=1) 
    app.missing_data_table_frame.grid_columnconfigure(0, weight=1)
    app.step_one_frame.grid_rowconfigure(5, weight=1)  
    app.step_one_frame.grid_columnconfigure(0, weight=1)
    
    #Create a description for the table 
    app.missing_data_table_description = customtkinter.CTkLabel(app.missing_data_table_frame, text="The SQs displayed below have missing metadata. Please update in Aconex.",
                                                    font=customtkinter.CTkFont(size=15))
    app.missing_data_table_description.grid(row=0, column=0, sticky='n')

    #Simplify column names and create the Treeview widget
    simplified_columns = [col.replace('?', '').replace('(', '').replace(')', '').replace('/', '_').replace(' ', '_') for col in dataframe.columns]
    app.missing_data_table = ttk.Treeview(app.missing_data_table_frame, columns=simplified_columns, show='headings')
    app.missing_data_table.grid(row=1, column=0, padx=20, pady=10, sticky='nsew')

    #Create and place a vertical scrollbar
    scrollbar = ttk.Scrollbar(app.missing_data_table_frame, orient="vertical", command=app.missing_data_table.yview)
    scrollbar.grid(row=1, column=1, sticky='ns')
    app.missing_data_table.configure(yscrollcommand=scrollbar.set)

    #Download Button for the Excel file.
    app.download_missing_metadata_file_button = customtkinter.CTkButton(app.missing_data_table_frame, text="Download Missing Metadata Excel File", command=download_missing_incorrect_data_table)
    app.download_missing_metadata_file_button.grid(row=2, column=0, padx=20, pady=10)

    #Dropdown menu to select the sheet.
    sheet_names = list(missing_incorrect_dfs.keys())
    SQs_With_Missing_Data = customtkinter.StringVar(value=initial_sheet_name)
    sheet_selector = customtkinter.CTkOptionMenu(app.missing_data_table_frame, values=sheet_names, variable=SQs_With_Missing_Data, command=update_table, corner_radius=0)
    sheet_selector.grid(row=2, column=0, sticky='wn', padx=20, pady=0)

    #Insert data into the table
    update_table(initial_sheet_name)

    if sheet_names:
        update_table(sheet_names[0])

def create_new_SQ_batch_table(app, excel_file_path):

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
        app.SQs_to_send_data_table.delete(*app.SQs_to_send_data_table.get_children())
        app.SQs_to_send_data_table['columns'] = simplified_columns

        #Create new columns based on the current dataframe.
        for col, simplified_col_name in zip(dataframe.columns, simplified_columns):
            app.SQs_to_send_data_table.heading(simplified_col_name, text=col)
            app.SQs_to_send_data_table.column(simplified_col_name, anchor="center")

        #Insert new data.
        for _, row in dataframe.iterrows():
            values = [row[col] for col in dataframe.columns]
            app.SQs_to_send_data_table.insert("", "end", values=values)

    #Create a frame for table and description
    app.SQs_to_send_data_table_frame = customtkinter.CTkFrame(app.step_one_frame, corner_radius=0)
    app.SQs_to_send_data_table_frame.grid(row=6, column=0, padx=20, pady=10, sticky='new')
    app.SQs_to_send_data_table_frame.grid_rowconfigure(1, weight=1) 
    app.SQs_to_send_data_table_frame.grid_columnconfigure(0, weight=1)
    app.step_one_frame.grid_rowconfigure(5, weight=1)  
    app.step_one_frame.grid_columnconfigure(0, weight=1)
    
    #Create a description for the table 
    app.SQs_to_send_data_table_description = customtkinter.CTkLabel(app.SQs_to_send_data_table_frame, text="The SQs displayed below will be sent to the City in this batch. Check to see if the metadata matches the PDF file.",
                                                    font=customtkinter.CTkFont(size=15))
    app.SQs_to_send_data_table_description.grid(row=0, column=0, sticky='n')

    #Simplify column names and create the Treeview widget
    simplified_columns = [col.replace('?', '').replace('(', '').replace(')', '').replace('/', '_').replace(' ', '_') for col in dataframe.columns]
    app.SQs_to_send_data_table = ttk.Treeview(app.SQs_to_send_data_table_frame, columns=simplified_columns, show='headings')
    app.SQs_to_send_data_table.grid(row=1, column=0, padx=20, pady=10, sticky='nsew')

    #Create and place a vertical scrollbar
    scrollbar = ttk.Scrollbar(app.SQs_to_send_data_table_frame, orient="vertical", command=app.SQs_to_send_data_table.yview)
    scrollbar.grid(row=1, column=1, sticky='ns')
    app.SQs_to_send_data_table.configure(yscrollcommand=scrollbar.set)

    #Download Button for the Excel file.
    app.download_missing_metadata_file_button = customtkinter.CTkButton(app.SQs_to_send_data_table_frame, text="Download SQ Checklist Excel File", command=download_SQs_to_send_file)
    app.download_missing_metadata_file_button.grid(row=2, column=0, padx=20, pady=10)

    #Dropdown menu to select the sheet.
    sheet_names_SQs_to_send = list(SQs_to_send_dfs.keys())
    SQs_to_send_Data = customtkinter.StringVar(value=initial_sheet_name_2)
    sheet_selector_SQs_to_send = customtkinter.CTkOptionMenu(app.SQs_to_send_data_table_frame, values=sheet_names_SQs_to_send, variable=SQs_to_send_Data, command=update_table_SQs_to_send, corner_radius=0)
    sheet_selector_SQs_to_send.grid(row=2, column=0, sticky='wn', padx=20, pady=0)

    #Insert data into the table
    update_table_SQs_to_send(initial_sheet_name_2)

    if sheet_names_SQs_to_send:
        update_table_SQs_to_send(sheet_names_SQs_to_send[0])

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

