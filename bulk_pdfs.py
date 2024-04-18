import os
import shutil
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from tkinterdnd2 import DND_FILES, TkinterDnD
import fitz  # PyMuPDF library, make sure it's installed
import threading
import signal


def signal_handler(sig, frame):
    sys.exit(0)
signal.signal(signal.SIGINT, signal_handler)

def close_app(event=None):
    root.destroy()
    
def check():
    root.after(50, check) # 50 stands for 50 ms.
    



def open_pdf_to_print(file_path):
    os.startfile(file_path, "print")

def setup_ui(root, output_directories):
    
    def log_message(message, status="info"):
        """
        Logs a message to the output_text widget.
        'status' can be 'info' for general information, 'success' for success messages, or 'error' for error messages.
        """
        # Colors for different statuses
        colors = {
            "info": "black",
            "success": "green",
            "error": "red"
        }
        color = colors.get(status, "black")

        # Insert the message
        output_text.config(state='normal')
        output_text.insert(tk.END, message + "\n", status)
        output_text.tag_config(status, foreground=color)
        output_text.config(state='disabled')
        output_text.yview(tk.END)  # Auto-scroll to the bottom
    
    def resize_window(files_list, root):
        longest_path = max((files_list.get(idx) for idx in range(files_list.size())), key=len, default='')
        estimated_width = min(16 * len(longest_path), root.winfo_screenwidth())  # Assume character width of 16 pixels
        new_width = max(1200, estimated_width)  # Ensure minimum width of 1200
        root.geometry(f'{new_width}x700')  # Adjust height as needed
    
    def select_image_file():
        filepath = filedialog.askopenfilename(
            title='Select Image File', 
            filetypes=[('Image Files', '*.jpg *.jpeg *.png *.bmp *.gif')]
        )
        if filepath:
            image_path_var.set(filepath)
            update_process_button_state(action_var, files_list, process_button)
            update_image_selection_visibility(action_var.get())
            image_file_selected_var.set(True)
        else:
            image_file_selected_var.set(False)
        update_image_selection_visibility(action_var.get())

            
    def update_image_selection_visibility(action):
        if action == "Insert Image":
            image_path_label.pack(in_=image_group_frame, fill=tk.X)
            select_image_button.pack(in_=image_group_frame, fill=tk.X)
            image_file_selected = image_file_selected_var.get()
            if image_file_selected == True:
                coord_frame.pack(in_=image_group_frame, fill=tk.X)
                
            else:
                coord_frame.pack_forget()

        elif action == "Extract First Pages":
            image_path_label.pack_forget()
            select_image_button.pack_forget()
            coord_frame.pack_forget()
        else:
            image_path_label.pack_forget()
            select_image_button.pack_forget()
            coord_frame.pack_forget()
            
    def change_output_directory():
        directory = os.path.abspath(filedialog.askdirectory())
        if directory:
            output_directory_var.set(directory)
            log_message(f"\n\n* * *\nOutput directories updated to: {directory}\n* * *\n\n")
            
            # Update the output_directories dictionary with new paths
            output_directories['outputs_dir'] = directory
            output_directories['backup_dir'] = os.path.join(directory, 'backed-up-originals')
            output_directories['image_inserted_dir'] = os.path.join(directory, 'Image-Inserted-On-Each')
            output_directories['first_pages_dir'] = os.path.join(directory, 'First-Pages')
            output_directories['first_pages_individual_dir'] = os.path.join(directory, 'First-Pages', 'First-Pages-Individual')
            output_directories['first_pages_combined_dir'] = os.path.join(directory, 'First-Pages', 'Combined')
            
            # Create the subdirectories
            os.makedirs(output_directories['backup_dir'], exist_ok=True)
            os.makedirs(output_directories['image_inserted_dir'], exist_ok=True)
            os.makedirs(output_directories['first_pages_individual_dir'], exist_ok=True)
            os.makedirs(output_directories['first_pages_combined_dir'], exist_ok=True)
                      
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
            process_button.config(state=tk.DISABLED, bg="grey")
        elif files_list.size() > 0 and action != "Select action...":
            process_button.config(state=tk.NORMAL, bg="green")
        else:
            process_button.config(state=tk.DISABLED, bg="grey")
            
    def delete_selected_items(event=None):
        # Get selected items, this returns a tuple of selected indices
        selected_items = files_list.curselection()

        # If nothing is selected, do nothing
        if not selected_items:
            return

        # Reverse the list of selected items, then delete from the end
        for i in reversed(selected_items):
            files_list.delete(i)
    
    # These are helper functions to update the UI and show error messages
    def update_ui_function(message):
        # Update the UI with the progress message
        # This could be adding to a status listbox or changing a label, for example
        pass
    
    
    def process_insert_image(files, image_path, output_directories, root):
        img_pixmap = fitz.Pixmap(image_path)
        
        # Retrieve coordinates from entry widgets
        x_start = int(x_coord.get())
        width = int(width_entry.get())
        y_start = int(y_coord.get())
        height = int(height_entry.get())
        
        # Calculate x_end and y_end based on start and width/height
        x_end = x_start + width
        y_end = y_start + height
        rect = fitz.Rect(x_start, y_start, x_end, y_end)
        
        log_message("\nAdding image to each PDF at the following coordinates:\n")
        log_message(f"\tWidth: From X_0: {x_start} to X_end: {width}")
        log_message(f"\tHeight: From Y_0 {y_start} to Y_end: {height}")
        
        for file_path in files:
            output_dir = output_directories["image_inserted_dir"]
            try:
                shutil.copy(file_path, output_dir)
            except Exception as e:
                error_message = "Failed to copy file for {}: {}".format(file_path, str(e))
                print(f"\n!\n! Error!\n!\n! {error_message}\n! {e}\n!\n")
                root.after(0, show_error_message, error_message)
            
            new_file_path = os.path.join(output_dir, os.path.basename(file_path))
            
            try:
                # Logic to insert the image into the PDF file goes here
                print(f"About to insert image at the top of file: [{file_path}]")
                add_image_to_pdf(new_file_path, rect, img_pixmap)
                # Use root.after to safely update the UI from the main thread
                root.after(0, update_ui_function, "Completed image insertion for {}".format(file_path))
            except Exception as e:
                error_message = "Failed to insert image for {}: {}".format(file_path, str(e))
                print(f"\n!\n! Error!\n!\n! {error_message}\n! {e}\n!\n")
                root.after(0, show_error_message, error_message)
        img_pixmap = None  # Release the pixmap
        
    
    def add_image_to_pdf(new_file_path, rect, img_pixmap):
        doc = fitz.open(new_file_path)
        first_page = doc[0]
        first_page.insert_image(rect, pixmap=img_pixmap)
        doc.save(new_file_path, incremental=True, encryption=fitz.PDF_ENCRYPT_KEEP)
        doc.close()
        log_message(f"Completed image insertion for: {new_file_path}")
    
    
    def process_extract_first_pages(files, output_directories, root):
        combined_pdf = fitz.open()  # Create a new PDF for combining first pages
        
        count = 0
        for file_path in files:
            try:
                # Logic to extract first pages goes here
                print(f"About to extract the first page for: [{file_path}]")
                doc = fitz.open(file_path)
                first_page = doc.load_page(0)
                output_pdf = fitz.open()
                output_pdf.insert_pdf(doc, from_page=0, to_page=0)
                individual_output_path = os.path.join(output_directories['first_pages_individual_dir'], 
                                                      os.path.basename(file_path))
                output_pdf.save(individual_output_path)
                count += 1
                log_message(f"{count}. Successfully extracted the 1st page of PDF to output dir: \n\t{individual_output_path}\n")
                output_pdf.close()
                
                # Add to combined PDF
                combined_pdf.insert_pdf(doc, from_page=0, to_page=0)
                
                doc.close()
                    
                # Use root.after to safely update the UI from the main thread
                root.after(0, update_ui_function, "Extracted first page for {}".format(file_path))
            
            except Exception as e:
                error_message = "Failed to extract first page for {}: {}".format(file_path, str(e))
                root.after(0, show_error_message, error_message)
        
        # Save the combined PDF
        combined_output_path = os.path.join(output_directories['first_pages_combined_dir'], 'combined_first_pages.pdf')
        try:
            combined_pdf.save(combined_output_path)
            log_message(f"Successfully combined each first-page into single PDF: {combined_output_path}", "success")
        except Exception as e:
            log_message(f"Failed to combined each first-page into single PDF {combined_output_path}: {str(e)}", "error")
        combined_pdf.close()

    
            
    def process_selected_action(action_var, files_list, image_path_var, output_directories, root):
        backup_originals(files_list, output_directories.get("backup_dir"))
        
        log_message("---------------------------------------------------------------------------")
        files = files_list.get(0, tk.END)
        action = action_var.get()
        if action == "Insert Image":
            image_path = image_path_var.get()
            if image_path:
                threading.Thread(
                    target=process_insert_image, 
                    args=(files, image_path, output_directories, root)
                ).start()
            else:
                messagebox.showerror("Error", "No image file selected.")
        elif action == "Extract First Pages":
            threading.Thread(
                target=process_extract_first_pages, 
                args=(files, output_directories, root)
            ).start()
    
    
    # setup_ui initialisation starts here:
    # =====================================
            
    root.title('PDF Processing Tool')
    root.geometry('1200x700')
    image_file_selected_var = tk.BooleanVar(value=False)
    
    # Configure the input-box (left) frame
    left_frame = tk.Frame(root, width=600, pady=20)  
    left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=False, padx=20, pady=20)
    left_frame.pack_propagate(False)
    
    files_list = tk.Listbox(left_frame, selectmode='extended', width=900)
    files_list.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
    
    add_button = tk.Button(left_frame, text="Add Files...",
                       font=('Helvetica', 14, 'bold'), bg='blue', fg='white',
                       command=lambda: add_files(files_list, action_var, process_button))
    add_button.pack(side=tk.BOTTOM, fill=tk.X)
    
    # Configure the fixed-width middle frame
    middle_frame = tk.Frame(root, width=200)  # Set a width that suits your needs
    middle_frame.pack(side=tk.LEFT, fill=tk.Y, pady=40)
    middle_frame.pack_propagate(False)
    
    # Configure the output-box (right) frame
    right_frame = tk.Frame(root, pady=20)  # Decreased maximum size
    right_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    output_text = tk.Text(right_frame, state='disabled', bg="white", font=('Helvetica', 10))
    output_text.pack(fill=tk.BOTH, padx=20, pady=20, expand=True)

    
    # Group related items in frames for better control
    action_group_frame = tk.Frame(middle_frame)
    action_group_frame.pack(fill=tk.X)  # Add padding as needed for group separation

    vspace_frame_1 = tk.Frame(middle_frame, height=50)
    vspace_frame_1.pack(fill=tk.X)

    image_group_frame = tk.Frame(middle_frame, height=300)
    
    
    image_group_frame.pack(fill=tk.X)  # Increase pady for more separation between groups
    image_group_frame.pack_propagate(False)
    
    vspace_frame_2 = tk.Frame(middle_frame, height=50)
    vspace_frame_2.pack(fill=tk.X)

    output_path_group_frame = tk.Frame(middle_frame)
    output_path_group_frame.pack(fill=tk.X, pady=20)  # Increase pady for more separation between groups

    process_list_group_frame = tk.Frame(middle_frame)
    process_list_group_frame.pack(side=tk.BOTTOM, fill=tk.X)  # Add padding as needed for group separation
    
    action_var = tk.StringVar(value="Select action...")
    actions = ["Extract First Pages", "Insert Image"]
    action_dropdown = tk.OptionMenu(root, action_var, *actions)
    action_dropdown.pack(in_=action_group_frame, side=tk.TOP, fill=tk.X)

    image_path_var = tk.StringVar()
    image_path_label = tk.Label(root, textvariable=image_path_var)
    image_path_label.pack_forget()  # Start hidden
    select_image_button = tk.Button(root, 
                                    text="Select Image...",
                                    font=('Helvetica', 12, 'bold'), bg='blue', fg='white',
                                    command=select_image_file)
    select_image_button.pack_forget()  # Start hidden

    # Main coordinates frame
    coord_frame = tk.Frame(image_group_frame)
    # coord_frame.pack(side=tk.TOP, padx=10, pady=10)

    # First row for X coordinates
    x_frame = tk.Frame(coord_frame)
    x_frame.pack(fill=tk.X)

    x_label = tk.Label(x_frame, text="X_0:")
    x_label.pack(side=tk.LEFT, padx=5)
    x_coord = tk.Entry(x_frame, width=5)
    x_coord.pack(side=tk.LEFT)
    x_coord.insert(0, '100')

    width_label = tk.Label(x_frame, text="X_end (Width):")
    width_label.pack(side=tk.LEFT, padx=5)
    width_entry = tk.Entry(x_frame, width=5)
    width_entry.pack(side=tk.LEFT)
    width_entry.insert(0, '300')

    # Spacer frame for vertical separation, if needed
    spacer_frame = tk.Frame(coord_frame, height=10)
    spacer_frame.pack(fill=tk.X)
    spacer_frame.pack_propagate(False)  # This ensures the frame keeps its size even if empty

    # Second row for Y coordinates
    y_frame = tk.Frame(coord_frame)
    y_frame.pack(fill=tk.X)

    y_label = tk.Label(y_frame, text="Y_0:")
    y_label.pack(side=tk.LEFT, padx=5)
    y_coord = tk.Entry(y_frame, width=5)
    y_coord.pack(side=tk.LEFT)
    y_coord.insert(0, '100')

    height_label = tk.Label(y_frame, text="Y_end (Height):")
    height_label.pack(side=tk.LEFT, padx=5)
    height_entry = tk.Entry(y_frame, width=5)
    height_entry.pack(side=tk.LEFT)
    height_entry.insert(0, '100')
    
    # Add a button to the middle frame to change the output directory
    output_directory_var = tk.StringVar()
    output_directory_var.set(output_directories["outputs_dir"]) 
    output_dir_label = tk.Label(root, textvariable=output_directory_var)
    output_dir_label.pack(in_=output_path_group_frame, fill=tk.X)
    change_output_dir_button = tk.Button(
        middle_frame, 
        text="Change Output Directory",
        command=change_output_directory
    )
    change_output_dir_button.pack(in_=output_path_group_frame, side=tk.TOP)
    
    process_button = tk.Button(
        root, 
        text="Process List", 
        state=tk.DISABLED,
        font=('Helvetica', 14, 'bold'), bg='green', fg='white',
        command=lambda: process_selected_action(action_var, files_list, image_path_var, output_directories, root)
    )
    process_button.config(state=tk.DISABLED, bg="grey")
    process_button.pack(in_=process_list_group_frame, fill=tk.X)

    files_list.drop_target_register(DND_FILES)
    files_list.dnd_bind('<<Drop>>', 
                        lambda event: drop(event, files_list, action_var, process_button, root))
    
    action_var.trace_add("write", lambda *args: update_process_button_state(action_var, files_list, process_button))
    action_var.trace_add("write", lambda *args: update_image_selection_visibility(action_var.get()))
    
    files_list.bind('<Delete>', delete_selected_items)
        
    
def setup_directories():
    base_dir = os.getcwd()
    backup_dir = os.path.join(base_dir, 'backed-up-originals')
    outputs_dir = os.path.join(base_dir, 'outputs')
    image_inserted_dir = os.path.join(outputs_dir, 'Image-Inserted-On-Each')
    first_pages_dir = os.path.join(outputs_dir, 'First-Pages', 'First-Pages-Individual')
    first_pages_individual_dir = os.path.join(outputs_dir, 'First-Pages', 'First-Pages-Individual')
    first_pages_combined_dir = os.path.join(outputs_dir, 'First-Pages', 'Combined')
    print("Creating directories...")
    os.makedirs(backup_dir, exist_ok=True)
    os.makedirs(image_inserted_dir, exist_ok=True)
    os.makedirs(first_pages_dir, exist_ok=True)
    os.makedirs(first_pages_combined_dir, exist_ok=True)
    print("done.\n")

    return {
        'outputs_dir': outputs_dir,
        'backup_dir': backup_dir,
        'image_inserted_dir': image_inserted_dir,
        'first_pages_dir': first_pages_dir,
        'first_pages_individual_dir': first_pages_individual_dir,
        'first_pages_combined_dir': first_pages_combined_dir
    }
    


def show_error_message(message):
    print("Error occurred: " + message, file=sys.stderr)

    # Safely handle GUI updates or closing
    if root and root.winfo_exists():  
        messagebox.showerror("Error", message)
        try:
            root.destroy()
            
        except Exception as e:
            print("Error while trying to close the application: " + str(e), file=sys.stderr)
    sys.exit(1)  # Ensure the application exits

def backup_originals(files_list, backup_dir):
    for file_path in files_list.get(0, tk.END):
        try:
            file_name = os.path.basename(file_path)
            dest_path = os.path.join(backup_dir, file_name)
            shutil.copy2(file_path, dest_path)
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while backing up file: {file_path}")
    
if __name__ == "__main__":
    try:
        output_directories = setup_directories()  
        root = TkinterDnD.Tk()  
        setup_ui(root, output_directories)  
        root.bind('<Control-c>', close_app)
        root.after(50, check)
        root.mainloop()
    
    except Exception as e:
        show_error_message("An unrecoverable error occurred: {}".format(e))
        raise
    
    
    
