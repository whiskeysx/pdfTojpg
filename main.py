import os
import threading
import zipfile
import rarfile
from tkinter import Tk, Label, Button, Entry, filedialog, messagebox, StringVar, OptionMenu, Checkbutton, BooleanVar
from tkinter import ttk
from pdf2image import convert_from_path
from pathlib import Path
import time
import shutil

# Define the global variable root
root = None

# Function to create output folders
def create_output_folders(pdf_files, output_base_folder):
    output_folders = []
    for pdf_file in pdf_files:
        folder_name = os.path.splitext(os.path.basename(pdf_file))[0]
        output_folder = os.path.join(output_base_folder, folder_name)
        os.makedirs(output_folder, exist_ok=True)
        output_folders.append((output_folder, folder_name))
    return output_folders

# Function to convert PDF to JPG in a separate thread
def convert_pdf_to_jpg(pdf_files, output_folders, poppler_path, progress_callback):
    try:
        total_files = len(pdf_files)
        start_time = time.time()
        
        for idx, (pdf_file, (output_folder, base_name)) in enumerate(zip(pdf_files, output_folders)):
            # Convert PDF pages to images
            pages = convert_from_path(pdf_file, 300, poppler_path=poppler_path)
            num_pages = len(pages)
            print(f"PDF {pdf_file} contains {num_pages} pages.")
            
            image_files = []  # List to keep track of created image files
            for i, page in enumerate(pages):
                output_file = os.path.join(output_folder, f"{base_name}_page_{i + 1}.jpg")
                page.save(output_file, "JPEG")
                image_files.append(output_file)
                print(f"Saved: {output_file}")

                # Update progress
                progress_message = f"Converting {os.path.basename(pdf_file)} - Page {i + 1} of {num_pages}"
                progress_callback((idx + 1) / total_files * 100, progress_message)

            # Update progress after each file
            progress_callback((idx + 1) / total_files * 100, f"Completed {os.path.basename(pdf_file)}")

            # Compress images
            zip_path = compress_images(image_files, compression_format.get(), chapter_structure_var.get(), output_folder, base_name)

            # Move the compressed file to the base output folder
            if zip_path:
                shutil.move(zip_path, os.path.join(output_base_folder, os.path.basename(zip_path)))

        elapsed_time = time.time() - start_time
        progress_message = f"Conversion and compression complete in {elapsed_time:.2f} seconds"
        messagebox.showinfo("Success", "All PDFs converted and compressed successfully!")
        progress_callback(100, progress_message)
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")
        progress_callback(0, "Error occurred")

def compress_images(image_files, file_type, chapter_structure, output_folder, base_name):
    if chapter_structure:
        # Create a subfolder for each compressed file if chapter structure is enabled
        compression_base_folder = os.path.join(output_folder, f"{base_name}_compressed")
        os.makedirs(compression_base_folder, exist_ok=True)
    else:
        compression_base_folder = output_folder

    zip_path = os.path.join(compression_base_folder, f"{base_name}.{file_type}")
    
    if file_type in ['zip', 'cbz']:
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            for image_file in image_files:
                zipf.write(image_file, os.path.basename(image_file))
    elif file_type == 'cbr':
        with rarfile.RarFile(zip_path, 'w') as rar:
            for image_file in image_files:
                rar.write(image_file, os.path.basename(image_file))
    
    return zip_path

def start_conversion():
    global root  # Declare root as global

    # Get the PDF paths and base output path from the entries
    pdf_files = pdf_path_entry.get().split(';')
    output_base_folder = output_path_entry.get()

    # Check if the paths are provided
    if not pdf_files or not output_base_folder:
        messagebox.showwarning("Warning", "Please provide both PDF path and output base folder.")
        return

    # Check if output base folder exists
    if not os.path.isdir(output_base_folder):
        messagebox.showwarning("Warning", "Output base folder does not exist.")
        return

    # Create output folders for each PDF file
    output_folders = create_output_folders(pdf_files, output_base_folder)

    # Disable buttons and start progress
    convert_button.config(state='disabled')
    progress_bar['value'] = 0
    status_label.config(text="Starting conversion...")
    root.update_idletasks()

    # Convert PDFs to JPGs in a separate thread
    threading.Thread(target=convert_pdf_to_jpg, args=(pdf_files, output_folders, poppler_path, update_progress)).start()

def browse_pdf_path():
    pdf_files = filedialog.askopenfilenames(title="Select PDF Files", filetypes=[("PDF files", "*.pdf")])
    if pdf_files:
        pdf_path_entry.delete(0, 'end')
        pdf_path_entry.insert(0, ';'.join(pdf_files))

def browse_output_path():
    output_folder = filedialog.askdirectory(title="Select Output Base Folder")
    if output_folder:
        output_path_entry.delete(0, 'end')
        output_path_entry.insert(0, output_folder)

def update_progress(value, message):
    progress_bar['value'] = value
    status_label.config(text=message)
    root.update_idletasks()

# Main GUI Window
def create_gui():
    global root, pdf_path_entry, output_path_entry, convert_button, progress_bar, status_label, compression_format, chapter_structure_var
    
    root = Tk()
    root.title("PDF to JPG Converter")
    root.geometry("600x450")

    # Add labels, input fields, and buttons to the window
    Label(root, text="PDF Path (separate multiple files with ';'):", font=("Helvetica", 12)).pack(pady=5)
    
    pdf_path_entry = Entry(root, width=70)
    pdf_path_entry.pack(pady=5)
    
    browse_pdf_button = Button(root, text="Browse PDFs", command=browse_pdf_path)
    browse_pdf_button.pack(pady=5)

    Label(root, text="Output Base Path:", font=("Helvetica", 12)).pack(pady=5)
    
    output_path_entry = Entry(root, width=70)
    output_path_entry.pack(pady=5)
    
    browse_output_button = Button(root, text="Browse Output Base Folder", command=browse_output_path)
    browse_output_button.pack(pady=5)

    # Compression options
    compression_format = StringVar(root)
    compression_format.set('zip')  # Default value

    Label(root, text="Select Compression Format:", font=("Helvetica", 12)).pack(pady=5)
    
    compression_menu = OptionMenu(root, compression_format, 'zip', 'cbz', 'cbr')
    compression_menu.pack(pady=5)

    # Chapter Structure checkbox
    chapter_structure_var = BooleanVar()
    chapter_structure_checkbox = Checkbutton(root, text="Chapter Structure", variable=chapter_structure_var, font=("Helvetica", 12))
    chapter_structure_checkbox.pack(pady=10)

    convert_button = Button(root, text="Convert PDFs to JPGs", command=start_conversion, padx=10, pady=5)
    convert_button.pack(pady=10)

    # Progress bar
    progress_bar = ttk.Progressbar(root, length=550, mode='determinate')
    progress_bar.pack(pady=10)

    # Status label
    status_label = Label(root, text="Waiting for user input...", font=("Helvetica", 10))
    status_label.pack(pady=10)

    # Start the GUI loop
    root.mainloop()

# Define Poppler path
poppler_path = r'C:\poppler\bin'  # Adjust this to your actual Poppler path

# Run the GUI
if __name__ == "__main__":
    create_gui()
