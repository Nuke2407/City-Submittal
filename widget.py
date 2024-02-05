import customtkinter
from tkinter import filedialog, messagebox, ttk
import os
from PIL import Image
import shutil

from excel_manipulation import Excel_Manipulation
from tables import *
from save_load_functionality import *

class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()
        
        self.init_ui()
        load_state(self)

    def init_ui(self):
        #Loads the images. 
        image_path = os.path.join(os.path.dirname(os.path.realpath(__file__)), "images")
        self.logo_image = customtkinter.CTkImage(light_image=Image.open(os.path.join(image_path, "Marigold_Logo_01_dark.png")),
                                                 dark_image=Image.open(os.path.join(image_path, "Marigold_Logo_01_light.png")), size=(179,44.5))
        self.step_one = customtkinter.CTkImage(light_image=Image.open(os.path.join(image_path, "one_dark.png")),
                                                 dark_image=Image.open(os.path.join(image_path, "one_light.png")), size=(20, 20))
        self.step_two = customtkinter.CTkImage(light_image=Image.open(os.path.join(image_path, "two_dark.png")),
                                                 dark_image=Image.open(os.path.join(image_path, "two_light.png")), size=(20, 20))
        icon_path = os.path.join(image_path, "cropped-Marigold-favicon-01.ico")
        
        #Set up the initial window.
        self.geometry("1130x1000")
        self.title("City Submittal")
        self.iconbitmap(icon_path)

        #Set grid layout 1x2.
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)
        
        #Create navigation frame on the left hand side. 
        self.navigation_frame = customtkinter.CTkFrame(self, corner_radius=0)
        self.navigation_frame.grid(row=0, column=0, sticky="nsew")
        self.navigation_frame.grid_rowconfigure(4, weight=1)

        self.navigation_frame_label = customtkinter.CTkLabel(self.navigation_frame, text="", image=self.logo_image)
        self.navigation_frame_label.grid(row=0, column=0, padx=20, pady=20)

        #Create two step buttons in the navigation frame. 
        self.step_one_button = customtkinter.CTkButton(self.navigation_frame, corner_radius=0, height=40, border_spacing=10, text="Step One",
                                                   fg_color="transparent", text_color=("gray10", "gray90"), hover_color=("gray70", "gray30"),
                                                   image=self.step_one, anchor="w", command=self.step_one_button_event)
        self.step_one_button.grid(row=1, column=0, sticky="ew")

        self.step_two_button = customtkinter.CTkButton(self.navigation_frame, corner_radius=0, height=40, border_spacing=10, text="Step Two",
                                                      fg_color="transparent", text_color=("gray10", "gray90"), hover_color=("gray70", "gray30"),
                                                      image=self.step_two, anchor="w", command=self.step_two_button_event)
        self.step_two_button.grid(row=2, column=0, sticky="ew")

        #Create a button to download the procedures 
        self.download_procedures_button = customtkinter.CTkButton(self.navigation_frame, text="Download Procedures", fg_color="green", command=self.download_procedures, state="disabled")
        self.download_procedures_button.grid(row=4, column=0, padx=20, pady=10)

        #Create an appearance mode dropdown where you can choose between dark and light modes. 
        self.appearance_mode_menu = customtkinter.CTkOptionMenu(self.navigation_frame, values=["Dark", "Light"],
                                                                command=self.change_appearance_mode_event)
        self.appearance_mode_menu.grid(row=6, column=0, padx=20, pady=20, sticky="s")

        #Reset Button.
        self.reset_button = customtkinter.CTkButton(self.navigation_frame, text="Reset", fg_color="red", command=self.reset_state)
        self.reset_button.grid(row=5, column=0, padx=20, pady=10)

        #Create step 1 frame
        self.step_one_frame = customtkinter.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.step_one_frame.grid_columnconfigure(0, weight=1)

        self.step_one_Title = customtkinter.CTkLabel(self.step_one_frame, text="City Submittal", font=customtkinter.CTkFont(size=30, weight="bold"))
        self.step_one_Title.grid(row=0, column=0, padx=20, pady=(20, 10))

        self.step_one_instructions = customtkinter.CTkLabel(self.step_one_frame, text="Rename and upload 'ExportDocs_Step_1.xls' and 'VLW-LOG-11000050-DC-0001_SQ_old.xlsx' files to the specified buttons below.",
                                                            font=customtkinter.CTkFont(size=15))
        self.step_one_instructions.grid(row=1, column=0, padx=20, pady=(20, 10))

        self.upload_Stage_1_button_exportdocs = customtkinter.CTkButton(self.step_one_frame, text="Upload The Export Document - ExportDocs_Step_1.xls", command=self.upload_exportdocs_step_1)
        self.upload_Stage_1_button_exportdocs.grid(row=2, column=0, padx=20, pady=10)

        self.upload_button_sqmetadata = customtkinter.CTkButton(self.step_one_frame, text="Upload The Previous SQ Log - VLW-LOG-11000050-DC-0001_SQ_old.xlsx", command=self.upload_sqmetadata)
        self.upload_button_sqmetadata.grid(row=3, column=0, padx=20, pady=10)

        self.process_files_button = customtkinter.CTkButton(self.step_one_frame, text="Process Files", command=self.process_files, state="disabled")
        self.process_files_button.grid(row=4, column=0, padx=20, pady=10)


        #Create step 2 frame.
        self.step_two_frame = customtkinter.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.step_two_frame.grid_columnconfigure(0, weight=1)

        self.step_two_instructions = customtkinter.CTkLabel(self.step_two_frame, text="Rename and upload the 'ExportDocs_Step_2.xls' export document with the missing/incorrect data from before updated in Aconex.", font=customtkinter.CTkFont(size=15))
        self.step_two_instructions.grid(row=0, column=0, padx=20, pady=(20,10))

        self.upload_Stage_2_button_exportdocs = customtkinter.CTkButton(self.step_two_frame, text="Please complete Step 1 before uploading the second export file.", command=self.upload_exportdocs_step_2, state="disabled")
        self.upload_Stage_2_button_exportdocs.grid(row=1, column=0, padx=20, pady=10)

        #Select default frame.
        self.select_frame_by_name("Step One")

        #initialize attributes to prevent errors on closing and reseting window.   
        self.missing_incorrect_data_table = None
        self.SQs_to_send_data_table = None
        self.missing_incorrect_data_table_stage_2 = None
        self.step_two_instructions = None
        self.city_submittal_data_table_stage_2 = None
        self.process_files_stage_2_button = None

        self.missing_incorrect_data_table_frame = None
        self.SQs_to_send_data_table_frame = None
        self.missing_incorrect_data_table_stage_2_frame = None
        self.city_submittal_data_table_stage_2_frame = None
        self.upload_download_buttons_frame = None
        
        #Override the window closing event.
        self.protocol("WM_DELETE_WINDOW",lambda: on_closing(self))
        
        #Create a save file if it does not exit. 
        if not os.path.exists("save_data"):  
            os.makedirs("save_data")

    def select_frame_by_name(self, name):
        #Set button color for selected button.
        self.step_one_button.configure(fg_color=("gray75", "gray25") if name == "Step One" else "transparent")
        self.step_two_button.configure(fg_color=("gray75", "gray25") if name == "Step Two" else "transparent")

        #Show selected frame.
        if name == "Step One":
            self.step_one_frame.grid(row=0, column=1, sticky="nsew")
        else:
            self.step_one_frame.grid_forget()
        if name == "Step Two":
            self.step_two_frame.grid(row=0, column=1, sticky="nsew")
        else:
            self.step_two_frame.grid_forget()

    def step_one_button_event(self):
        self.select_frame_by_name("Step One")

    def step_two_button_event(self):
        self.select_frame_by_name("Step Two")

    def change_appearance_mode_event(self, new_appearance_mode):
        customtkinter.set_appearance_mode(new_appearance_mode)
        with open("save_data/appearance_mode.txt","w") as file:
            file.write(new_appearance_mode)

    def reset_state(self):
        #Delete the state file if it exists
        state_file = "save_data/app_state.json"
        table_file = "save_data/table_data.json"
        if os.path.exists(state_file):
            os.remove(state_file)
        if os.path.exists(table_file):
            os.remove(table_file)
        
        #Reset UI elements to default
        self.upload_Stage_1_button_exportdocs.configure(state="normal", text="Upload The Export Document - ExportDocs_Step_1.xls")
        self.upload_button_sqmetadata.configure(state="normal", text="Upload The Previous SQ Log - VLW-LOG-11000050-DC-0001_SQ_old.xls")
        self.process_files_button.configure(state="disabled")

        self.upload_Stage_2_button_exportdocs.configure(state="disabled", text="Please complete Step 1 before uploading the second export file.")

        #Delete the table an buttons
        if self.missing_incorrect_data_table_frame is not None: 
            self.missing_incorrect_data_table_frame.destroy()
            self.missing_incorrect_data_table_frame = None
            self.missing_incorrect_data_table = None
        if self.SQs_to_send_data_table_frame is not None: 
            self.SQs_to_send_data_table_frame.destroy()
            self.SQs_to_send_data_table_frame = None
            self.SQs_to_send_data_table = None
        if self.missing_incorrect_data_table_stage_2_frame is not None: 
            self.missing_incorrect_data_table_stage_2_frame.destroy()
            self.missing_incorrect_data_table_stage_2_frame = None
            self.missing_incorrect_data_table_stage_2 = None
        if self.step_two_instructions is not None:
            self.step_two_instructions.destroy()
            self.step_two_instructions = None
        if self.upload_download_buttons_frame is not None:
            self.upload_download_buttons_frame.destroy()
            self.upload_download_buttons_frame = None
        if self.process_files_stage_2_button is not None:
            self.process_files_stage_2_button.destroy()
            self.process_files_stage_2_button = None
        if self.city_submittal_data_table_stage_2_frame is not None:
            self.city_submittal_data_table_stage_2_frame.destroy()
            self.city_submittal_data_table_stage_2_frame = None
            self.city_submittal_data_table_stage_2 = None

        # Check if the file exists
        if os.path.exists('stage_one_documents/sq_removal_excel.xlsx'):
            try:
                # Attempt to delete the file
                os.remove('stage_one_documents/sq_removal_excel.xlsx')
            except Exception as e:
                # If an error occurs during deletion, show an error messagebox
                messagebox.showerror("Error", f"An error occurred: {e}")

    #STEP 1 MEATHODS
    def upload_exportdocs_step_1(self):
        filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xls")])
        if filepath:
            if os.path.basename(filepath) == "ExportDocs_Step_1.xls":
                #Correct file selected, move it to the stage_one_documents folder
                destination_folder = os.path.join(os.path.dirname(os.path.realpath(__file__)), "stage_one_documents")
                shutil.copy(filepath, os.path.join(destination_folder, "ExportDocs_Step_1.xls"))
                self.upload_Stage_1_button_exportdocs.configure(state="disabled", text="Export Document Uploaded")
            else:
                #Incorrect file name, show an error message
                messagebox.showerror("Error", "Please select a file named 'ExportDocs_Step_1.xlsx'.")
        self.check_both_files_uploaded()

    def upload_sqmetadata(self):
        
        filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if filepath:
            if os.path.basename(filepath) == "VLW-LOG-11000050-DC-0001_SQ_old.xlsx":
                #Correct file selected, move it to the stage_one_documents folder
                destination_folder = os.path.join(os.path.dirname(os.path.realpath(__file__)), "stage_one_documents")
                shutil.copy(filepath, os.path.join(destination_folder, "VLW-LOG-11000050-DC-0001_SQ_old.xlsx"))
                self.upload_button_sqmetadata.configure(state="disabled", text="Previous SQ Log Uploaded")
            else:
                #Incorrect file name, show an error message
                messagebox.showerror("Error", "Please select a file named 'VLW-LOG-11000050-DC-0001_SQ_old.xlsx'.")
        self.check_both_files_uploaded()

    def check_both_files_uploaded(self):
        if not self.upload_Stage_1_button_exportdocs.cget('state') == 'disabled' or not self.upload_button_sqmetadata.cget('state') == 'disabled':
            self.process_files_button.configure(state="disabled")
        else:
            self.process_files_button.configure(state="normal")

    def process_files(self):
        Excel_Manipulation().stage_1()

        #Clear existing table if it exists
        if self.missing_incorrect_data_table_frame is not None:
            self.missing_incorrect_data_table_frame.destroy()
        if self.SQs_to_send_data_table_frame is not None:
            self.SQs_to_send_data_table_frame.destroy()

        #Create a table and display the DataFrame
        create_missing_incorrect_data_table_stage_1(self, 'data/missing_incorrect_metadata_file.xlsx')
        create_new_SQ_batch_table(self, 'data/new_city_sub_SQs.xlsx')

        self.upload_Stage_2_button_exportdocs.configure(state="normal", text="Upload The Second Export Document - ExportDocs_Step_2.xls")


    #STEP 2 FUNCTIONS
    def upload_exportdocs_step_2(self):
        filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xls")])
        if filepath:
            if os.path.basename(filepath) == "ExportDocs_Step_2.xls":
                #Correct file selected, move it to the stage_one_documents folder
                destination_folder = os.path.join(os.path.dirname(os.path.realpath(__file__)), "stage_one_documents")
                shutil.copy(filepath, os.path.join(destination_folder, "ExportDocs_Step_2.xls"))

                Excel_Manipulation().stage_2_part_1()

                #Clear existing table if it exists
                if self.missing_incorrect_data_table_stage_2_frame is not None:
                    self.missing_incorrect_data_table_stage_2_frame.destroy()

                create_missing_incorrect_data_table_stage_2(self, 'data/missing_incorrect_metadata_file_2.xlsx')

                self.sq_removal_ui_creation()
            else:
                #Incorrect file name, show an error message
                messagebox.showerror("Error", "Please select a file named 'ExportDocs_Step_2.xlsx'.")

    def sq_removal_ui_creation(self):
        self.step_two_instructions = customtkinter.CTkLabel(self.step_two_frame, text="Below are download and upload links for a file where you can input the SQs that are on hold and those that have to be superceeded.", font=customtkinter.CTkFont(size=15))
        self.step_two_instructions.grid(row=3, column=0, padx=20, pady=(20,10))

        self.upload_download_buttons_frame = customtkinter.CTkFrame(self.step_two_frame, corner_radius=0, fg_color="transparent")
        self.upload_download_buttons_frame.grid(row=4, column=0, padx=20, pady=10, sticky='new')
        self.upload_download_buttons_frame.grid_rowconfigure(1, weight=1) 
        self.upload_download_buttons_frame.grid_columnconfigure(0, weight=1)
        self.upload_download_buttons_frame.grid_columnconfigure(1, weight=1)

        self.download_on_hold_superseded_file_button = customtkinter.CTkButton(self.upload_download_buttons_frame, text="Download On Hold/Superseded SQ Format", command=self.download_sq_removal_excel_fomat)
        self.download_on_hold_superseded_file_button.grid(row=1, column=0)

        self.upload_on_hold_superseded_file_button = customtkinter.CTkButton(self.upload_download_buttons_frame, text="Upload sq_removal_excel.xlsx File", command=self.upload_sq_removal_excel_file)
        self.upload_on_hold_superseded_file_button.grid(row=1, column=1)

        self.process_files_stage_2_button = customtkinter.CTkButton(self.step_two_frame, text="Process Files", command=self.process_files_stage_2)
        self.process_files_stage_2_button.grid(row=5, column=0)
    
    def download_sq_removal_excel_fomat(self):
        try:
            #Assuming the file is saved in the 'data' directory
            source_file = 'data/permanent/sq_removal_excel_fomat.xlsx'
            if os.path.exists(source_file):
                # Ask the user where to save the file
                filetypes = [('Excel File', '*.xlsx'), ('All Files', '*.*')]
                dest_file = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=filetypes, title="Save File As")
                # If the user selects a location, copy the file
                if dest_file:
                    shutil.copy(source_file, dest_file)             
            else:
                messagebox.showerror("Download Failed", "The source file does not exist.")
        
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")
        
    def upload_sq_removal_excel_file(self):
        filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if filepath:
            if os.path.basename(filepath) == "sq_removal_excel.xlsx":
                #Correct file selected, move it to the stage_one_documents folder
                destination_folder = os.path.join(os.path.dirname(os.path.realpath(__file__)), "stage_one_documents")
                shutil.copy(filepath, os.path.join(destination_folder, "sq_removal_excel.xlsx"))
            else:
                #Incorrect file name, show an error message
                messagebox.showerror("Error", "Please select a file named 'sq_removal_excel.xlsx'.")

    def process_files_stage_2(self):
        Excel_Manipulation().stage_2_part_2()

        #Clear existing table if it exists
        if self.city_submittal_data_table_stage_2_frame is not None:
            self.city_submittal_data_table_stage_2_frame.destroy()

        create_city_submittal_final_document_data_table_stage_2(self, 'data/City Submittal.xlsx')

    def download_procedures(self):
        try:
            #Assuming the file is saved in the 'data' directory
            source_file = 'data/permanent/procedures.pdf'
            if os.path.exists(source_file):
                # Ask the user where to save the file
                filetypes = [('Excel File', '*.pdf'), ('All Files', '*.*')]
                dest_file = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=filetypes, title="Save File As")
                # If the user selects a location, copy the file
                if dest_file:
                    shutil.copy(source_file, dest_file)             
            else:
                messagebox.showerror("Download Failed", "The source file does not exist.")
        
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

if __name__ == "__main__":
    #Ensure the save_data directory exists
    if not os.path.exists("save_data"):  
        os.makedirs("save_data")

    app = App()
    app.mainloop()