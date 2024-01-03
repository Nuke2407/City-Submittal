import customtkinter
from tkinter import filedialog, messagebox, ttk
import os
from PIL import Image
import shutil
import json 
import pandas

from stage_1 import Stage_1

class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()
        
        self.init_ui()
        self.load_state()

    def init_ui(self):
        #Loads the images. 
        image_path = os.path.join(os.path.dirname(os.path.realpath(__file__)), "images")
        self.logo_image = customtkinter.CTkImage(light_image=Image.open(os.path.join(image_path, "Marigold_Logo_01_dark.png")),
                                                 dark_image=Image.open(os.path.join(image_path, "Marigold_Logo_01_light.png")), size=(179,44.5))
        self.step_one = customtkinter.CTkImage(light_image=Image.open(os.path.join(image_path, "one_dark.png")),
                                                 dark_image=Image.open(os.path.join(image_path, "one_light.png")), size=(20, 20))
        self.step_two = customtkinter.CTkImage(light_image=Image.open(os.path.join(image_path, "two_dark.png")),
                                                 dark_image=Image.open(os.path.join(image_path, "two_light.png")), size=(20, 20))
        self.step_three = customtkinter.CTkImage(light_image=Image.open(os.path.join(image_path, "three_dark.png")),
                                                     dark_image=Image.open(os.path.join(image_path, "three_light.png")), size=(20, 20))
        icon_path = os.path.join(image_path, "cropped-Marigold-favicon-01.ico")
        
        #Set up the initial window.
        self.geometry("1100x800")
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

        #Create three buttons in the navigation frame. 
        self.step_one_button = customtkinter.CTkButton(self.navigation_frame, corner_radius=0, height=40, border_spacing=10, text="Step One",
                                                   fg_color="transparent", text_color=("gray10", "gray90"), hover_color=("gray70", "gray30"),
                                                   image=self.step_one, anchor="w", command=self.step_one_button_event)
        self.step_one_button.grid(row=1, column=0, sticky="ew")

        self.step_two_button = customtkinter.CTkButton(self.navigation_frame, corner_radius=0, height=40, border_spacing=10, text="Step Two",
                                                      fg_color="transparent", text_color=("gray10", "gray90"), hover_color=("gray70", "gray30"),
                                                      image=self.step_two, anchor="w", command=self.step_two_button_event)
        self.step_two_button.grid(row=2, column=0, sticky="ew")

        self.step_three_button = customtkinter.CTkButton(self.navigation_frame, corner_radius=0, height=40, border_spacing=10, text="Step Three",
                                                      fg_color="transparent", text_color=("gray10", "gray90"), hover_color=("gray70", "gray30"),
                                                      image=self.step_three, anchor="w", command=self.step_three_button_event)
        self.step_three_button.grid(row=3, column=0, sticky="ew")

        #Create an appearance mode dropdown where you can choose between dark and light modes. 
        self.appearance_mode_menu = customtkinter.CTkOptionMenu(self.navigation_frame, values=["Light", "Dark", "System"],
                                                                command=self.change_appearance_mode_event)
        self.appearance_mode_menu.grid(row=6, column=0, padx=20, pady=20, sticky="s")

        #Create home frame.
        self.step_one = customtkinter.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.step_one.grid_columnconfigure(0, weight=1)
        self.step_one_Title = customtkinter.CTkLabel(self.step_one, text="City Submittal", font=customtkinter.CTkFont(size=30, weight="bold"))
        self.step_one_Title.grid(row=0, column=0, padx=20, pady=(20, 10))
        self.step_one_Instructions = customtkinter.CTkLabel(self.step_one, text="Re-label and upload the 'ExportDocs_Stage_1' and 'SQ_metadata_closed_template' files into the appropriate sections below.",
                                                            font=customtkinter.CTkFont(size=15))
        self.step_one_Instructions.grid(row=1, column=0, padx=20, pady=(20, 10))

        self.upload_button_exportdocs = customtkinter.CTkButton(self.step_one, text="Upload The Export Document - ExportDocs_Stage_1.xls", command=self.upload_exportdocs)
        self.upload_button_exportdocs.grid(row=2, column=0, padx=20, pady=10)
        self.upload_button_sqmetadata = customtkinter.CTkButton(self.step_one, text="Upload The Previous SQ Log - VLW-LOG-11000050-DC-0001_SQ_old.xls", command=self.upload_sqmetadata)
        self.upload_button_sqmetadata.grid(row=3, column=0, padx=20, pady=10)

        self.process_files_button = customtkinter.CTkButton(self.step_one, text="Process Files", command=self.process_files, state="disabled")
        self.process_files_button.grid(row=4, column=0, padx=20, pady=10)

        #Reset Button.
        self.reset_button = customtkinter.CTkButton(self.navigation_frame, text="Reset", command=self.reset_state)
        self.reset_button.grid(row=5, column=0, padx=20, pady=10)

        #Create second frame.
        self.step_two = customtkinter.CTkFrame(self, corner_radius=0, fg_color="transparent")

        #Create third frame.
        self.step_three = customtkinter.CTkFrame(self, corner_radius=0, fg_color="transparent")

        #Select default frame.
        self.select_frame_by_name("Step One")

        #initialize attributes to prevent errors on closing and reseting window.   
        self.table_frame = None
        self.missing_metadata_file = None

        #Override the window closing event.
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def select_frame_by_name(self, name):
        #Set button color for selected button.
        self.step_one_button.configure(fg_color=("gray75", "gray25") if name == "Step One" else "transparent")
        self.step_two_button.configure(fg_color=("gray75", "gray25") if name == "Step Two" else "transparent")
        self.step_three_button.configure(fg_color=("gray75", "gray25") if name == "Step Three" else "transparent")

        #Show selected frame.
        if name == "Step One":
            self.step_one.grid(row=0, column=1, sticky="nsew")
        else:
            self.step_one.grid_forget()
        if name == "Step Two":
            self.step_two.grid(row=0, column=1, sticky="nsew")
        else:
            self.step_two.grid_forget()
        if name == "Step Three":
            self.step_three.grid(row=0, column=1, sticky="nsew")
        else:
            self.step_three.grid_forget()

    #Functions responsible for changing the frame when a new frame is selected
    def step_one_button_event(self):
        self.select_frame_by_name("Step One")

    def step_two_button_event(self):
        self.select_frame_by_name("Step Two")

    def step_three_button_event(self):
        self.select_frame_by_name("Step Three")

    def change_appearance_mode_event(self, new_appearance_mode):
        customtkinter.set_appearance_mode(new_appearance_mode)
        with open("save_data/appearance_mode.txt","w") as file:
            file.write(new_appearance_mode)

    def upload_exportdocs(self):
        filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xls")])
        if filepath:
            if os.path.basename(filepath) == "ExportDocs_Stage_1.xls":
                #Correct file selected, move it to the stage_one_documents folder
                destination_folder = os.path.join(os.path.dirname(os.path.realpath(__file__)), "stage_one_documents")
                shutil.move(filepath, os.path.join(destination_folder, "ExportDocs_Stage_1.xls"))
                self.upload_button_exportdocs.configure(state="disabled", text="Export Document Uploaded")
            else:
                #Incorrect file name, show an error message
                messagebox.showerror("Error", "Please select a file named 'ExportDocs_Stage_1.xlsx'.")
        self.check_both_files_uploaded()

    def upload_sqmetadata(self):
        filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xls")])
        if filepath:
            if os.path.basename(filepath) == "VLW-LOG-11000050-DC-0001_SQ_old.xls":
                #Correct file selected, move it to the stage_one_documents folder
                destination_folder = os.path.join(os.path.dirname(os.path.realpath(__file__)), "stage_one_documents")
                shutil.move(filepath, os.path.join(destination_folder, "VLW-LOG-11000050-DC-0001_SQ_old.xls"))
                self.upload_button_sqmetadata.configure(state="disabled", text="Previous SQ Log Uploaded")
            else:
                #Incorrect file name, show an error message
                messagebox.showerror("Error", "Please select a file named 'VLW-LOG-11000050-DC-0001_SQ_old.xls'.")
        self.check_both_files_uploaded()

    def check_both_files_uploaded(self):
        if not self.upload_button_exportdocs.cget('state') == 'disabled' or not self.upload_button_sqmetadata.cget('state') == 'disabled':
            self.process_files_button.configure(state="disabled")
        else:
            self.process_files_button.configure(state="normal")

    def process_files(self):
        self.missing_metadata_file = Stage_1().SQs_missing_data()

        #Clear existing table if it exists
        if hasattr(self, 'table'):
            self.table.destroy()

        #Create a table and display the DataFrame
        self.create_table(self.missing_metadata_file)

    def create_table(self, dataframe):
        #Create a frame for table and description
        self.table_frame = customtkinter.CTkFrame(self.step_one, corner_radius=0)
        self.table_frame.grid(row=5, column=0, padx=20, pady=10, sticky='new')
        self.table_frame.grid_rowconfigure(1, weight=1) 
        self.table_frame.grid_columnconfigure(0, weight=1)
        self.step_one.grid_rowconfigure(5, weight=1)  
        self.step_one.grid_columnconfigure(0, weight=1)
        
        #Create a description for the table 
        self.table_description = customtkinter.CTkLabel(self.table_frame, text="The SQs displayed below have missing metadata. Please update in Aconex.",
                                                        font=customtkinter.CTkFont(size=15))
        self.table_description.grid(row=0, column=0, sticky='n')

        #Simplify column names and create the Treeview widget
        simplified_columns = [col.replace('?', '').replace('(', '').replace(')', '').replace('/', '_').replace(' ', '_') for col in dataframe.columns]
        self.table = ttk.Treeview(self.table_frame, columns=simplified_columns, show='headings')
        self.table.grid(row=1, column=0, padx=20, pady=10, sticky='nsew')

        #Define column headings and configure columns
        for col_name, simplified_col_name in zip(dataframe.columns, simplified_columns):
            self.table.heading(simplified_col_name, text=col_name)
            self.table.column(simplified_col_name, anchor="center")

        #Insert data into the table
        for _, row in dataframe.iterrows():
            values = [row[col] for col in dataframe.columns]
            self.table.insert("", "end", values=values)

        #Create and place a vertical scrollbar
        scrollbar = ttk.Scrollbar(self.table_frame, orient="vertical", command=self.table.yview)
        scrollbar.grid(row=1, column=1, sticky='ns')
        self.table.configure(yscrollcommand=scrollbar.set)

        self.download_missing_metadata_file_button = customtkinter.CTkButton(self.table_frame, text="Download Missing Metadata Excel File", command=self.download_missing_metadata_file)
        self.download_missing_metadata_file_button.grid(row=6, column=0, padx=20, pady=10)

    def download_missing_metadata_file(self):
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


    def load_state(self):
        try:
            with open("save_data/app_state.json", "r") as file:
                state = json.load(file)

                # Restore step
                self.select_frame_by_name(state.get("current_step", "Step One"))

                #Restore uploaded file states
                if state.get("exportdocs_uploaded", False):
                    self.upload_button_exportdocs.configure(state="disabled", text="Export Document Uploaded")

                if state.get("sqmetadata_uploaded", False):
                    self.upload_button_sqmetadata.configure(state="disabled", text="Previous SQ Log Uploaded")

                #If the table was displayed, recreate it (assuming you have the data available)
                if state.get("table_displayed", False):
                    with open("save_data/table_data.json", "r") as data_file:
                        table_data_json = data_file.read()
                        table_data = pandas.read_json(table_data_json, orient='records', lines=True)
                        self.create_table(table_data)

        except FileNotFoundError as e:
            if str(e).find('table_data.json') != -1:
                print("table_data.json not found. A new table will need to be created.")
            else:
                # Reset to default state if the state file is not found
                self.select_frame_by_name("Step One")
                self.upload_button_exportdocs.configure(state="normal", text="Upload The Export Document - ExportDocs_Stage_1.xls")
                self.upload_button_sqmetadata.configure(state="normal", text="Upload The Previous SQ Log - VLW-LOG-11000050-DC-0001_SQ_old.xls")
                self.process_files_button.configure(state="disabled")
                # Any other UI elements that need to be reset should be handled here

        except Exception as e:
            print("An error occurred:", e)


        try: 
            with open("save_data/appearance_mode.txt", "r") as file:
                mode = file.read().strip()
                customtkinter.set_appearance_mode(mode)
                self.appearance_mode_menu.set(mode)

        except FileNotFoundError:
            customtkinter.set_appearance_mode("Dark")
            self.appearance_mode_menu.set("Dark")

        except Exception as e:
            print("An error occurred:", e)

    def save_state(self, step_name):
        #Ensure that the 'table' attribute exists and is a widget before checking if it is mapped
        table_displayed = hasattr(self, 'table') and isinstance(self.table, ttk.Treeview) and self.table.winfo_exists() and self.table.winfo_ismapped()
        state = {
            "current_step": step_name,
            "exportdocs_uploaded": self.upload_button_exportdocs.cget("state") == "disabled",
            "sqmetadata_uploaded": self.upload_button_sqmetadata.cget("state") == "disabled",
            "table_displayed": table_displayed}
        
        with open("save_data/app_state.json", "w") as file:
            json.dump(state, file)

    def save_table_data(self, dataframe):
        table_data = dataframe.to_json(orient='records', lines=True)
        with open("save_data/table_data.json", "w") as file:
            file.write(table_data)

    def reset_state(self):
        #Delete the state file if it exists
        state_file = "save_data/app_state.json"
        table_file = "save_data/table_data.json"
        if os.path.exists(state_file):
            os.remove(state_file)
        if os.path.exists(table_file):
            os.remove(table_file)
        
        #Reset UI elements to default
        self.upload_button_exportdocs.configure(state="normal", text="Upload The Export Document - ExportDocs_Stage_1.xls")
        self.upload_button_sqmetadata.configure(state="normal", text="Upload The Previous SQ Log - VLW-LOG-11000050-DC-0001_SQ_old.xls")
        self.process_files_button.configure(state="disabled")

        #Delete the table
        self.table_frame.destroy()

        #Call load_state to reset the UI
        self.load_state()

    def on_closing(self):
        #Determine the current step and pass it to save_state before closing
        current_step = "Step One"  # default value
        if self.step_one.winfo_ismapped():
            current_step = "Step One"
        elif self.step_two.winfo_ismapped():
            current_step = "Step Two"
        elif self.step_three.winfo_ismapped():
            current_step = "Step Three"

        #Save state before attempting to save table data
        self.save_state(current_step)
        
        #Only attempt to save table data if the table exists and is a widget
        if hasattr(self, 'missing_metadata_file') and self.missing_metadata_file is not None and hasattr(self, 'table') and isinstance(self.table, ttk.Treeview):
            self.save_table_data(self.missing_metadata_file)
        
        self.destroy()

    #Outside of your App class, make sure to create the directory if it doesn't exist
    if not os.path.exists("save_data"):  
        os.makedirs("save_data")

    def data_for_table_is_available(self):
        return os.path.exists("save_data/table_data.json")

if __name__ == "__main__":
    #Ensure the save_data directory exists
    if not os.path.exists("save_data"):  
        os.makedirs("save_data")

    app = App()
    app.mainloop()