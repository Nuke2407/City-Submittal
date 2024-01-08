import customtkinter
from tkinter import filedialog, messagebox, ttk
import os
import shutil

def create_incorrect_data_table(app, dataframe):
    #Create a frame for table and description
    app.incorrect_data_table_frame = customtkinter.CTkFrame(app.step_one_frame, corner_radius=0)
    app.incorrect_data_table_frame.grid(row=5, column=0, padx=20, pady=10, sticky='new')
    app.incorrect_data_table_frame.grid_rowconfigure(1, weight=1) 
    app.incorrect_data_table_frame.grid_columnconfigure(0, weight=1)
    app.step_one_frame.grid_rowconfigure(5, weight=1)  
    app.step_one_frame.grid_columnconfigure(0, weight=1)
    
    #Create a description for the table 
    app.table_description = customtkinter.CTkLabel(app.incorrect_data_table_frame, text="The SQs displayed below have Incorrectly filled out metadata. Please update in Aconex.",
                                                    font=customtkinter.CTkFont(size=15))
    app.table_description.grid(row=0, column=0, sticky='n')

    #Simplify column names and create the Treeview widget
    simplified_columns = [col.replace('?', '').replace('(', '').replace(')', '').replace('/', '_').replace(' ', '_') for col in dataframe.columns]
    app.incorrect_data_table = ttk.Treeview(app.incorrect_data_table_frame, columns=simplified_columns, show='headings')
    app.incorrect_data_table.grid(row=1, column=0, padx=20, pady=10, sticky='nsew')

    #Define column headings and configure columns
    for col_name, simplified_col_name in zip(dataframe.columns, simplified_columns):
        app.incorrect_data_table.heading(simplified_col_name, text=col_name)
        app.incorrect_data_table.column(simplified_col_name, anchor="center")

    #Insert data into the table
    for _, row in dataframe.iterrows():
        values = [row[col] for col in dataframe.columns]
        app.incorrect_data_table.insert("", "end", values=values)

    #Create and place a vertical scrollbar
    scrollbar = ttk.Scrollbar(app.incorrect_data_table_frame, orient="vertical", command=app.incorrect_data_table.yview)
    scrollbar.grid(row=1, column=1, sticky='ns')
    app.incorrect_data_table.configure(yscrollcommand=scrollbar.set)

    app.download_missing_metadata_file_button = customtkinter.CTkButton(app.incorrect_data_table_frame, text="Download Missing Metadata Excel File", command=download_incorrect_data_table)
    app.download_missing_metadata_file_button.grid(row=6, column=0, padx=20, pady=10)

def download_incorrect_data_table():
    try:
        #Assuming the file is saved in the 'data' directory
        source_file = 'data/missing_metadata_file.xlsx'
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

def create_missing_data_table(app, dataframe):
    #Create a frame for table and description
    app.missing_data_table_frame = customtkinter.CTkFrame(app.step_one_frame, corner_radius=0)
    app.missing_data_table_frame.grid(row=6, column=0, padx=20, pady=10, sticky='new')
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

    #Define column headings and configure columns
    for col_name, simplified_col_name in zip(dataframe.columns, simplified_columns):
        app.missing_data_table.heading(simplified_col_name, text=col_name)
        app.missing_data_table.column(simplified_col_name, anchor="center")

    #Insert data into the table
    for _, row in dataframe.iterrows():
        values = [row[col] for col in dataframe.columns]
        app.missing_data_table.insert("", "end", values=values)

    #Create and place a vertical scrollbar
    scrollbar = ttk.Scrollbar(app.missing_data_table_frame, orient="vertical", command=app.missing_data_table.yview)
    scrollbar.grid(row=1, column=1, sticky='ns')
    app.missing_data_table.configure(yscrollcommand=scrollbar.set)

    app.download_missing_metadata_file_button = customtkinter.CTkButton(app.missing_data_table_frame, text="Download Missing Metadata Excel File", command=download_missing_metadata_file)
    app.download_missing_metadata_file_button.grid(row=6, column=0, padx=20, pady=10)

def download_missing_metadata_file():
    try:
        #Assuming the file is saved in the 'data' directory
        source_file = 'data/missing_metadata_file.xlsx'
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

