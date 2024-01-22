import customtkinter
from tkinter import ttk
import json 

from tables import *

def load_state(app_instance):
    try:
        with open("save_data/app_instance_state.json", "r") as file:
            state = json.load(file)

            # Restore step
            app_instance.select_frame_by_name(state.get("current_step", "Step One"))

            #Restore uploaded file states
            if state.get("exportdocs_uploaded", False):
                app_instance.upload_Stage_1_button_exportdocs.configure(state="disabled", text="Export Document Uploaded")

            if state.get("sqmetadata_uploaded", False):
                app_instance.upload_button_sqmetadata.configure(state="disabled", text="Previous SQ Log Uploaded")

            #If the tables were displayed, recreate it (assuming you have the data available)
            if state.get("table_displayed", False):
                create_missing_incorrect_data_table_stage_1(app_instance, 'data/missing_incorrect_metadata_file.xlsx')
                create_new_SQ_batch_table(app_instance, 'data/new_city_sub_SQs.xlsx')

            #If the stage 2 table was displayed, recreate it (assuming you have the data available)
            if state.get("table_displayed_stage_2", False):
                create_missing_incorrect_data_table_stage_2(app_instance, 'data/missing_incorrect_metadata_file_2.xlsx')

    except FileNotFoundError as e:
        if str(e).find('table_data.json') != -1:
            print("table_data.json not found. A new table will need to be created.")
        else:
            # Reset to default state if the state file is not found
            app_instance.select_frame_by_name("Step One")
            app_instance.upload_Stage_1_button_exportdocs.configure(state="normal", text="Upload The Export Document - ExportDocs_Stage_1.xls")
            app_instance.upload_button_sqmetadata.configure(state="normal", text="Upload The Previous SQ Log - VLW-LOG-11000050-DC-0001_SQ_old.xls")
            app_instance.process_files_button.configure(state="disabled")

    except Exception as e:
        print("An error occurred:", e)

    try: 
        with open("save_data/app_instanceearance_mode.txt", "r") as file:
            mode = file.read().strip()
            customtkinter.set_app_instanceearance_mode(mode)

    except FileNotFoundError:
        customtkinter.set_appearance_mode("Dark")

    except Exception as e:
        print("An error occurred:", e)

def on_closing(app_instance):
    #Determine the current step and pass it to save_state before closing
    current_step = "Step One"  # default value
    if app_instance.step_one_frame.winfo_ismapped():
        current_step = "Step One"
    elif app_instance.step_two_frame.winfo_ismapped():
        current_step = "Step Two"
    elif app_instance.step_three_frame.winfo_ismapped():
        current_step = "Step Three"

    
    save_state(app_instance, current_step)

    app_instance.destroy()

def save_state(app_instance, step_name):
    #Ensure that the 'table' attribute exists and is a widget before checking if it is mapp_instanceed
    table_missing_incorrect_data_table = app_instance.missing_incorrect_data_table is not None and isinstance(app_instance.missing_incorrect_data_table, ttk.Treeview) 
    table_SQs_to_send_data_table = app_instance.SQs_to_send_data_table is not None and isinstance(app_instance.SQs_to_send_data_table, ttk.Treeview)
    table_missing_incorrect_data_table_stage_2 = app_instance.missing_incorrect_data_table_stage_2 is not None and isinstance(app_instance.missing_incorrect_data_table_stage_2, ttk.Treeview)

    table_displayed = table_missing_incorrect_data_table and table_SQs_to_send_data_table
    table_displayed_stage_2 = table_missing_incorrect_data_table_stage_2

    state = {
        "current_step": step_name,
        "exportdocs_uploaded": app_instance.upload_Stage_1_button_exportdocs.cget("state") == "disabled",
        "sqmetadata_uploaded": app_instance.upload_button_sqmetadata.cget("state") == "disabled",
        "table_displayed": table_displayed,
        "tabel_displayed_stage_2": table_displayed_stage_2}
    
    with open("save_data/app_instance_state.json", "w") as file:
        json.dump(state, file)



