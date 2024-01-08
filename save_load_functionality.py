import customtkinter
from tkinter import ttk
import os
from PIL import Image
import json 
import pandas

from tables import *

def load_state(app):
    try:
        with open("save_data/app_state.json", "r") as file:
            state = json.load(file)

            # Restore step
            app.select_frame_by_name(state.get("current_step", "Step One"))

            #Restore uploaded file states
            if state.get("exportdocs_uploaded", False):
                app.upload_button_exportdocs.configure(state="disabled", text="Export Document Uploaded")

            if state.get("sqmetadata_uploaded", False):
                app.upload_button_sqmetadata.configure(state="disabled", text="Previous SQ Log Uploaded")

            #If the table was displayed, recreate it (assuming you have the data available)
            if state.get("table_displayed", False):
                with open("save_data/table_data.json", "r") as data_file:
                    table_data_json = data_file.read()
                    table_data = pandas.read_json(table_data_json, orient='records', lines=True)
                    create_missing_data_table(app, table_data)

    except FileNotFoundError as e:
        if str(e).find('table_data.json') != -1:
            print("table_data.json not found. A new table will need to be created.")
        else:
            # Reset to default state if the state file is not found
            app.select_frame_by_name("Step One")
            app.upload_button_exportdocs.configure(state="normal", text="Upload The Export Document - ExportDocs_Stage_1.xls")
            app.upload_button_sqmetadata.configure(state="normal", text="Upload The Previous SQ Log - VLW-LOG-11000050-DC-0001_SQ_old.xls")
            app.process_files_button.configure(state="disabled")
            # Any other UI elements that need to be reset should be handled here

    except Exception as e:
        print("An error occurred:", e)


    try: 
        with open("save_data/appearance_mode.txt", "r") as file:
            mode = file.read().strip()
            customtkinter.set_appearance_mode(mode)
            app.appearance_mode_menu.set(mode)

    except FileNotFoundError:
        customtkinter.set_appearance_mode("Dark")
        app.appearance_mode_menu.set("Dark")

    except Exception as e:
        print("An error occurred:", e)

def on_closing(app):
    #Determine the current step and pass it to save_state before closing
    current_step = "Step One"  # default value
    if app.step_one_frame.winfo_ismapped():
        current_step = "Step One"
    elif app.step_two_frame.winfo_ismapped():
        current_step = "Step Two"
    elif app.step_three_frame.winfo_ismapped():
        current_step = "Step Three"

    #Save state before attempting to save table data
    save_state(app, current_step)
    
    #Prepare a dictionary to store dataframes
    table_to_save = {}
    #Only attempt to save tables data if the table exists and is a widget

    if app.missing_data_table is not None and isinstance(app.missing_data_table, ttk.Treeview):
        table_to_save['missing_metadata_file'] = app.missing_metadata_file
    if app.incorrect_data_table is not None and isinstance(app.incorrect_data_table, ttk.Treeview):
        table_to_save['incorrect_data_file'] = app.incorrect_data_file
    
    if table_to_save:
        app.save_table_data(table_to_save)

    app.destroy()

def save_state(app, step_name):
    #Ensure that the 'table' attribute exists and is a widget before checking if it is mapped
    table_missing_metadata_displayed = app.missing_data_table is not None and isinstance(app.missing_data_table, ttk.Treeview) 
    tabel_incorrect_data_file_displayed = app.incorrect_data_table is not None and isinstance(app.incorrect_data_table, ttk.Treeview)

    table_displayed = table_missing_metadata_displayed or tabel_incorrect_data_file_displayed

    state = {
        "current_step": step_name,
        "exportdocs_uploaded": app.upload_button_exportdocs.cget("state") == "disabled",
        "sqmetadata_uploaded": app.upload_button_sqmetadata.cget("state") == "disabled",
        "table_displayed": table_displayed}
    
    with open("save_data/app_state.json", "w") as file:
        json.dump(state, file)

def save_table_data(app, table_dict):
    table_json = {key: df.to_json(orient='records', lines=True) for key, df in table_dict.items()}
    with open("save_data/table_data.json", "w") as file:
        json.dump(table_json, file)


