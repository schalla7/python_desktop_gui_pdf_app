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
    
    # def toggle_print_options():
    #     # This function will run whenever 'print_now_var' changes
    #     should_print = print_now_var.get()
    #     if should_print:
    #         # If 'Also print the output now' is checked, show the other options and set them to True
    #         print_individual_checkbox.pack(side=tk.TOP, pady=5, fill=tk.X)
    #         print_combined_checkbox.pack(side=tk.TOP, pady=5, fill=tk.X)
    #         print_individual_var.set(True)
    #         print_combined_var.set(True)
    #     else:
    #         # If 'Also print the output now' is unchecked, hide the other options and set them to False
    #         print_individual_checkbox.pack_forget()
    #         print_combined_checkbox.pack_forget()
    #         print_individual_var.set(False)
    #         print_combined_var.set(False)
    
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
            
    def update_image_selection_visibility(action):
        if action == "Insert Image":
            # select_image_button.pack_forget()  # Start hidden
    
            # process_button.config(state=tk.DISABLED, bg="grey", width=20)
            process_button.pack_forget()
            
    
            image_path_label.pack()
            select_image_button.pack()
            process_button.pack(padx=20, pady=20)
            # Hide the printing options
            # print_now_checkbox.pack_forget()
            # print_individual_checkbox.pack_forget()
            # print_combined_checkbox.pack_forget()
            # print_now_var.set(False)
            # print_individual_var.set(False)
            # print_combined_var.set(False)
            # Ensure the dependent checkboxes are also hidden
            # toggle_print_options()
            
        elif action == "Extract First Pages":
            image_path_label.pack_forget()
            select_image_button.pack_forget()
            # Show the printing options
            # print_now_checkbox.pack(side=tk.TOP, pady=5, fill=tk.X)
            # print_now_checkbox.pack(side=tk.TOP, fill=tk.X, padx=20)
    
            # print_individual_checkbox.pack(side=tk.TOP, pady=5, fill=tk.X)
            # print_combined_checkbox.pack(side=tk.TOP, pady=5, fill=tk.X)
            # Show or hide the dependent checkboxes based on the current state
            # toggle_print_options()
            
        else:
            image_path_label.pack_forget()
            select_image_button.pack_forget()
            # print_now_checkbox.pack_forget()
            # print_individual_checkbox.pack_forget()
            # print_combined_checkbox.pack_forget()
            # Ensure the dependent checkboxes are also hidden
            # toggle_print_options()
            
    def add_files(files_list, action_var, process_button):
        try:
            filenames = filedialog.askopenfilenames(title='Select PDF files', filetypes=[('PDF files', '*.pdf')])
            for filename in filenames:
                files_list.insert(tk.END, filename.strip('{}'))
            update_process_button_state(action_var, files_list, process_button)
        except KeyboardInterrupt:
            print("User interruption received. Exiting...")
            sys.exit(0)
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
        except KeyboardInterrupt:
            print("User interruption received. Exiting...")
            sys.exit(0)
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
        
        for file_path in files:
            output_dir = output_directories["image_inserted_dir"]
            
            try:
                shutil.copy(file_path, output_dir)
            except KeyboardInterrupt:
                print("User interruption received. Exiting...")
                sys.exit(0)
            except Exception as e:
                error_message = "Failed to copy file for {}: {}".format(file_path, str(e))
                print(f"\n!\n! Error!\n!\n! {error_message}\n! {e}\n!\n")
                root.after(0, show_error_message, error_message)
            
            new_file_path = os.path.join(output_dir, os.path.basename(file_path))
            
            try:
                # Logic to insert the image into the PDF file goes here
                print(f"About to insert image at the top of file: [{file_path}]")
                add_image_to_pdf(new_file_path, img_pixmap)
                # Use root.after to safely update the UI from the main thread
                root.after(0, update_ui_function, "Completed image insertion for {}".format(file_path))
            except KeyboardInterrupt:
                print("User interruption received. Exiting...")
                sys.exit(0)
            except Exception as e:
                error_message = "Failed to insert image for {}: {}".format(file_path, str(e))
                print(f"\n!\n! Error!\n!\n! {error_message}\n! {e}\n!\n")
                root.after(0, show_error_message, error_message)
        img_pixmap = None  # Release the pixmap
        
    
    def add_image_to_pdf(new_file_path, img_pixmap):
        doc = fitz.open(new_file_path)
        rect = fitz.Rect(50, 50, 150, 150)
        first_page = doc[0]
        first_page.insert_image(rect, pixmap=img_pixmap)
        doc.save(new_file_path, incremental=True, encryption=fitz.PDF_ENCRYPT_KEEP)
        doc.close()
        log_message(f"Completed: {new_file_path}")
    
    
    def process_extract_first_pages(files, output_directories, root):
        combined_pdf = fitz.open()  # Create a new PDF for combining first pages
        
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
                output_pdf.close()
                
                # Add to combined PDF
                combined_pdf.insert_pdf(doc, from_page=0, to_page=0)
                
                doc.close()
                    
                # Use root.after to safely update the UI from the main thread
                root.after(0, update_ui_function, "Extracted first page for {}".format(file_path))
            
            except KeyboardInterrupt:
                print("User interruption received. Exiting...")
                sys.exit(0)
                
            except Exception as e:
                error_message = "Failed to extract first page for {}: {}".format(file_path, str(e))
                root.after(0, show_error_message, error_message)
        
        # Save the combined PDF
        combined_output_path = os.path.join(output_directories['first_pages_combined_dir'], 'combined_first_pages.pdf')
        try:
            combined_pdf.save(combined_output_path)
            log_message(f"Successfully combined each first-page into single PDF: {combined_output_path}", "success")
        except KeyboardInterrupt:
            print("User interruption received. Exiting...")
            sys.exit(0)
        except Exception as e:
            log_message(f"Failed to combined each first-page into single PDF {combined_output_path}: {str(e)}", "error")
        combined_pdf.close()

        # Optional print of the combined PDF
        # if print_now and print_combined:
        #     open_pdf_to_print(combined_output_path)
    
            
    def process_selected_action(action_var, files_list, image_path_var, output_directories, root):
        backup_originals(files_list, output_directories.get("backup_dir"))
        
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
            
    root.title('PDF Processing Tool')
    root.geometry('1200x700')
    
    # Configure the input-box (left) frame
    left_frame = tk.Frame(root)  # Increased minimum size
    left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)
    files_list = tk.Listbox(left_frame, selectmode='extended')
    files_list.pack(fill=tk.BOTH, expand=True)
    

    # Configure the output-box (right) frame
    right_frame = tk.Frame(root)  # Decreased maximum size
    right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=10, pady=10)
    output_text = tk.Text(right_frame, state='disabled', bg="white", font=('Helvetica', 10))
    output_text.pack(fill=tk.BOTH, expand=True)

    action_var = tk.StringVar(value="Select action...")
    actions = ["Extract First Pages", "Insert Image"]
    action_dropdown = tk.OptionMenu(root, action_var, *actions)
    action_dropdown.pack(side=tk.TOP, fill=tk.X, padx=20, pady=20)

    add_button = tk.Button(left_frame, text="Add Files...",
                       font=('Helvetica', 14, 'bold'), bg='blue', fg='white',
                       command=lambda: add_files(files_list, action_var, process_button))
    add_button.config(width=20)
    add_button.pack(side=tk.BOTTOM, fill=tk.X, padx=20, pady=20)

    image_path_var = tk.StringVar()
    image_path_label = tk.Label(root, textvariable=image_path_var)
    image_path_label.pack_forget()  # Start hidden
    select_image_button = tk.Button(root, text="Select Image...", command=select_image_file)
    select_image_button.pack_forget()  # Start hidden
    
    process_button = tk.Button(
        root, 
        text="Process List", 
        state=tk.DISABLED,
        font=('Helvetica', 14, 'bold'), bg='green', fg='white',
        command=lambda: process_selected_action(action_var, files_list, image_path_var, output_directories, root)
    )
    process_button.config(state=tk.DISABLED, bg="grey", width=20)
    process_button.pack(padx=20, pady=20)

    files_list.drop_target_register(DND_FILES)
    files_list.dnd_bind('<<Drop>>', 
                        lambda event: drop(event, files_list, action_var, process_button, root))
    
    action_var.trace_add("write", lambda *args: update_process_button_state(action_var, files_list, process_button))
    action_var.trace_add("write", lambda *args: update_image_selection_visibility(action_var.get()))
    
    # print_now_var = tk.BooleanVar(value=True)
    # print_individual_var = tk.BooleanVar(value=True)
    # print_combined_var = tk.BooleanVar(value=True)

    # print_now_checkbox = tk.Checkbutton(root, text="Also print the output now", variable=print_now_var)
    # print_individual_checkbox = tk.Checkbutton(root, text="Print each first page separately", variable=print_individual_var)
    # print_combined_checkbox = tk.Checkbutton(root, text="Print combination of first pages", variable=print_combined_var)
    
    # Bind the 'toggle_print_options' function to changes in 'print_now_var'
    # print_now_var.trace_add('write', lambda *args: toggle_print_options())
    
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
        'backup_dir': backup_dir,
        'image_inserted_dir': image_inserted_dir,
        'first_pages_dir': first_pages_dir,
        'first_pages_individual_dir': first_pages_individual_dir,
        'first_pages_combined_dir': first_pages_combined_dir
    }

def show_error_message(message):
    print("Error occurred: " + message, file=sys.stderr)

    # Safely handle GUI updates or closing
    if root and root.winfo_exists():  # Check if root exists and is in a valid state
        messagebox.showerror("Error", message)
        try:
            root.destroy()
        except KeyboardInterrupt:
            print("User interruption received. Exiting...")
            sys.exit(0)
        except Exception as e:
            print("Error while trying to close the application: " + str(e), file=sys.stderr)
    sys.exit(1)  # Ensure the application exits

def backup_originals(files_list, backup_dir):
    for file_path in files_list.get(0, tk.END):
        try:
            file_name = os.path.basename(file_path)
            dest_path = os.path.join(backup_dir, file_name)
            shutil.copy2(file_path, dest_path)
        except KeyboardInterrupt:
            print("User interruption received. Exiting...")
            sys.exit(0)
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while backing up file: {file_path}")
    
if __name__ == "__main__":
    try:
        output_directories = setup_directories()  
        root = TkinterDnD.Tk()  
        setup_ui(root, output_directories)  
        root.mainloop()
    
    except KeyboardInterrupt:
        print("User interruption received. Exiting...")
        sys.exit(0)
        
    except Exception as e:
        show_error_message("An unrecoverable error occurred: {}".format(e))
        raise
    
