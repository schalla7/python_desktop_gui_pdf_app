import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from tkinterdnd2 import DND_FILES, TkinterDnD
import fitz  # PyMuPDF library, make sure it's installed
import threading



def open_pdf_to_print(file_path):
    os.startfile(file_path, "print")

def setup_ui(root, output_directories):
    
    def toggle_print_options():
        # This function will run whenever 'print_now_var' changes
        should_print = print_now_var.get()
        if should_print:
            # If 'Also print the output now' is checked, show the other options and set them to True
            print_individual_checkbox.pack(side=tk.TOP, pady=5, fill=tk.X)
            print_combined_checkbox.pack(side=tk.TOP, pady=5, fill=tk.X)
            print_individual_var.set(True)
            print_combined_var.set(True)
        else:
            # If 'Also print the output now' is unchecked, hide the other options and set them to False
            print_individual_checkbox.pack_forget()
            print_combined_checkbox.pack_forget()
            print_individual_var.set(False)
            print_combined_var.set(False)
    
    def resize_window(files_list, root):
        longest_path = max((files_list.get(idx) for idx in range(files_list.size())), key=len, default='')
        estimated_width = min(16 * len(longest_path), root.winfo_screenwidth())  # Assume character width of 16 pixels
        new_width = max(800, estimated_width)  # Ensure minimum width of 800
        root.geometry(f'{new_width}x600')  # Adjust height as needed
        
    
    def select_image_file():
        filepath = filedialog.askopenfilename(
            title='Select Image File', 
            filetypes=[('Image Files', '*.jpg *.jpeg *.png *.bmp *.gif')]
        )
        if filepath:
            image_path_var.set(filepath)
            update_process_button_state(action_var, files_list, process_button)
            
    def update_image_selection_visibility(action):
        if action == "Insert Image":
            image_path_label.pack()
            select_image_button.pack()
            # Hide the printing options
            print_now_checkbox.pack_forget()
            print_individual_checkbox.pack_forget()
            print_combined_checkbox.pack_forget()
            print_now_var.set(False)
            print_individual_var.set(False)
            print_combined_var.set(False)
            # Ensure the dependent checkboxes are also hidden
            toggle_print_options()
            
        elif action == "Extract First Pages":
            image_path_label.pack_forget()
            select_image_button.pack_forget()
            # Show the printing options
            print_now_checkbox.pack(side=tk.TOP, pady=5, fill=tk.X)
            print_individual_checkbox.pack(side=tk.TOP, pady=5, fill=tk.X)
            print_combined_checkbox.pack(side=tk.TOP, pady=5, fill=tk.X)
            # Show or hide the dependent checkboxes based on the current state
            toggle_print_options()
            
        else:
            image_path_label.pack_forget()
            select_image_button.pack_forget()
            print_now_checkbox.pack_forget()
            print_individual_checkbox.pack_forget()
            print_combined_checkbox.pack_forget()
            # Ensure the dependent checkboxes are also hidden
            toggle_print_options()
            
    def add_files(files_list, action_var, process_button):
        try:
            filenames = filedialog.askopenfilenames(title='Select PDF files', filetypes=[('PDF files', '*.pdf')])
            for filename in filenames:
                files_list.insert(tk.END, filename.strip('{}'))
            update_process_button_state(action_var, files_list, process_button)
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while adding files: {e}")
            raise

    def drop(event, files_list, action_var, process_button, root):
        try:
            files = event.widget.tk.splitlist(event.data)
            for file in files:
                files_list.insert("end", file.strip('{}'))
            update_process_button_state(action_var, files_list, process_button)
            resize_window(files_list, root)
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred during drop: {e}")
            raise
            
    def update_process_button_state(action_var, files_list, process_button):
        action = action_var.get()
        valid_image_selected = bool(image_path_var.get())
        
        if action == "Insert Image" and not valid_image_selected:
            process_button.config(state=tk.DISABLED)
        elif files_list.size() > 0 and action != "Select action...":
            process_button.config(state=tk.NORMAL)
        else:
            process_button.config(state=tk.DISABLED)
            
    def process_insert_image(files, image_path, output_dir, root):
        for file_path in files:
            try:
                # Logic to insert the image into the PDF file goes here
                print(f"About to insert image at the top of file: [{file_path}]")
                add_image_to_pdf(file_path, image_path, output_dir)
                # Use root.after to safely update the UI from the main thread
                root.after(0, update_ui_function, "Completed image insertion for {}".format(file_path))
            except Exception as e:
                error_message = "Failed to insert image for {}: {}".format(file_path, str(e))
                root.after(0, show_error_message, error_message)

    def process_extract_first_pages(files, output_directories, root):
        for file_path in files:
            try:
                # Logic to extract first pages goes here
                # You would pass the actual print_now_var.get(), etc. to this function as needed
                print(f"About to extract the first page for: [{file_path}]")
                extract_first_pages(file_path, output_directories['first_pages_dir'], True)
                # Use root.after to safely update the UI from the main thread
                root.after(0, update_ui_function, "Extracted first page for {}".format(file_path))
            except Exception as e:
                error_message = "Failed to extract first page for {}: {}".format(file_path, str(e))
                root.after(0, show_error_message, error_message)
    
    # These are helper functions to update the UI and show error messages
    def update_ui_function(message):
        # Update the UI with the progress message
        # This could be adding to a status listbox or changing a label, for example
        pass
    
    
        
    def add_image_to_pdf(file_path, img_obj, output_dir, update_status):
        doc = fitz.open(file_path)
        rect = fitz.Rect(0, 0, 50, 50)
        first_page = doc[0]
        first_page.insert_image(rect, stream=img_obj)
        new_file_path = os.path.join(output_dir, os.path.basename(file_path))
        doc.save(new_file_path, incremental=True, encryption=fitz.PDF_ENCRYPT_KEEP)
        doc.close()
        update_status(f"Completed: {file_path}")
        
    def extract_first_pages(file_path, output_dir, print_now):
        doc = fitz.open(file_path)
        first_page = doc.load_page(0)
        output_pdf = fitz.open()
        output_pdf.insert_pdf(doc, from_page=0, to_page=0)
        output_path = os.path.join(output_dir, os.path.basename(file_path))
        output_pdf.save(output_path)
        output_pdf.close()
        if print_now:
            open_pdf_to_print(output_path)
            # output_pdf.print()  # Simplified, adjust based on how you want to handle printing
        doc.close()
            
    def process_selected_action(action_var, files_list, image_path_var, output_directories, root):
        action = action_var.get()
        files = files_list.get(0, tk.END)
        
        if action == "Insert Image":
            image_path = image_path_var.get()
            if image_path:
                threading.Thread(
                    target=process_insert_image, 
                    args=(files, image_path, output_directories['image_inserted_dir'], root)
                ).start()
            else:
                messagebox.showerror("Error", "No image file selected.")
        elif action == "Extract First Pages":
            threading.Thread(
                target=process_extract_first_pages, 
                args=(files, output_directories, root)
            ).start()
        
            
    root.title('PDF Processing Tool')
    root.geometry('800x600')

    left_frame = tk.Frame(root)
    right_frame = tk.Frame(root)
    left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)
    right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=10, pady=10)

    files_list = tk.Listbox(left_frame)
    files_list.pack(fill=tk.BOTH, expand=True)

    output_list = tk.Listbox(right_frame)
    output_list.pack(fill=tk.BOTH, expand=True)

    action_var = tk.StringVar(value="Select action...")
    actions = ["Extract First Pages", "Insert Image"]
    action_dropdown = tk.OptionMenu(root, action_var, *actions)
    action_dropdown.pack(side=tk.TOP, pady=20)

    add_button = tk.Button(left_frame, text="Add PDFs...", 
                            command=lambda: add_files(files_list, action_var, process_button))
    add_button.pack()

    image_path_var = tk.StringVar()
    image_path_label = tk.Label(root, textvariable=image_path_var)
    image_path_label.pack_forget()  # Start hidden
    select_image_button = tk.Button(root, text="Select Image...", command=select_image_file)
    select_image_button.pack_forget()  # Start hidden
    
    process_button = tk.Button(
        root, 
        text="Process List", 
        state=tk.DISABLED, 
        command=lambda: process_selected_action(action_var, files_list, image_path_var, output_directories, root)
    )

    process_button.pack(side=tk.BOTTOM, pady=10)

    files_list.drop_target_register(DND_FILES)
    files_list.dnd_bind('<<Drop>>', 
                        lambda event: drop(event, files_list, action_var, process_button, root))
    
    action_var.trace_add("write", lambda *args: update_process_button_state(action_var, files_list, process_button))
    action_var.trace_add("write", lambda *args: update_image_selection_visibility(action_var.get()))
    
    print_now_var = tk.BooleanVar(value=True)
    print_individual_var = tk.BooleanVar(value=True)
    print_combined_var = tk.BooleanVar(value=True)

    print_now_checkbox = tk.Checkbutton(root, text="Also print the output now", variable=print_now_var)
    print_individual_checkbox = tk.Checkbutton(root, text="Print each first page separately", variable=print_individual_var)
    print_combined_checkbox = tk.Checkbutton(root, text="Print combination of first pages", variable=print_combined_var)
    # Bind the 'toggle_print_options' function to changes in 'print_now_var'
    print_now_var.trace_add('write', lambda *args: toggle_print_options())
        
    
def setup_directories():
    base_dir = os.getcwd()
    backup_dir = os.path.join(base_dir, 'backed-up-originals')
    outputs_dir = os.path.join(base_dir, 'outputs')
    image_inserted_dir = os.path.join(outputs_dir, 'Image-Inserted-On-Each')
    first_pages_dir = os.path.join(outputs_dir, 'First-Pages', 'First-Pages-Individual')
    combined_dir = os.path.join(outputs_dir, 'First-Pages', 'Combined')

    os.makedirs(backup_dir, exist_ok=True)
    os.makedirs(image_inserted_dir, exist_ok=True)
    os.makedirs(first_pages_dir, exist_ok=True)
    os.makedirs(combined_dir, exist_ok=True)

    return {
        'backup_dir': backup_dir,
        'image_inserted_dir': image_inserted_dir,
        'first_pages_dir': first_pages_dir,
        'combined_dir': combined_dir
    }

def show_error_message(message):
    print("Error occurred: " + message, file=sys.stderr)

    # Safely handle GUI updates or closing
    if root and root.winfo_exists():  # Check if root exists and is in a valid state
        messagebox.showerror("Error", message)
        try:
            root.destroy()
        except Exception as e:
            print("Error while trying to close the application: " + str(e), file=sys.stderr)
    sys.exit(1)  # Ensure the application exits

if __name__ == "__main__":
    try:
        output_directories = setup_directories()  
        root = TkinterDnD.Tk()  
        setup_ui(root, output_directories)  
        root.mainloop()
        
    except Exception as e:
        show_error_message("An unrecoverable error occurred: {}".format(e))
        raise
    
