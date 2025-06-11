import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os

# Assuming extractor.py and main.py are in the same directory or PYTHONPATH
from extractor import extraer_por_codigo
from main import save_extracted_data

# Imports for DOCX generation
from docx import Document
from llenar_docx import cargar_json, detectar_campos, reemplazar_campos, completar_listado


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Data Extractor and Doc Generator")

        # PDF Processing related variables
        self.pdf_path = tk.StringVar()
        self.subject_code_var = tk.StringVar()
        self.current_json_path = None # Stores path to the JSON from extraction step

        # DOCX Generation related variables
        self.docx_template_path = tk.StringVar()
        self.docx_dynamic_fields_entries = {} # Stores Entry StringVars for template fields

        # General status
        self.status_var = tk.StringVar()

        self.root.geometry("800x850") # Adjusted height slightly

        # Main frame
        main_frame = ttk.Frame(self.root, padding="10 10 10 10") # Uniform padding
        main_frame.pack(fill=tk.BOTH, expand=True)

        # PDF Processing Frame
        pdf_frame = ttk.LabelFrame(main_frame, text="1. PDF Processing", padding="10 10 10 10")
        pdf_frame.pack(fill=tk.X, pady=10) # Increased pady

        # PDF File Selection
        select_pdf_button = ttk.Button(pdf_frame, text="Select PDF", command=self.select_pdf_file)
        select_pdf_button.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.pdf_path_label = ttk.Label(pdf_frame, text="No PDF selected")
        self.pdf_path_label.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        # Subject Code Input
        subject_code_label = ttk.Label(pdf_frame, text="Subject Code:")
        subject_code_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")
        subject_code_entry = ttk.Entry(pdf_frame, textvariable=self.subject_code_var, width=30)
        subject_code_entry.grid(row=1, column=1, padx=5, pady=5, sticky="ew")

        # Extract Data Button
        extract_button = ttk.Button(pdf_frame, text="Extract Data from PDF", command=self.process_extraction)
        extract_button.grid(row=2, column=0, columnspan=2, pady=(10,5)) # pady top and bottom

        pdf_frame.grid_columnconfigure(1, weight=1) # Ensure label/entry expands

        # Results Display Frame (for Treeview)
        results_frame = ttk.LabelFrame(main_frame, text="2. Extracted Data Viewer", padding="10 10 10 10")
        results_frame.pack(fill=tk.BOTH, expand=True, pady=10) # Increased pady

        self.results_treeview = ttk.Treeview(results_frame, columns=("Orden", "Apellido y Nombre", "Puntos"), show="headings")
        self.results_treeview.heading("Orden", text="Orden")
        self.results_treeview.heading("Apellido y Nombre", text="Apellido y Nombre")
        self.results_treeview.heading("Puntos", text="Puntos")

        self.results_treeview.column("Orden", width=60, anchor=tk.CENTER, stretch=tk.NO)
        self.results_treeview.column("Apellido y Nombre", width=350, stretch=tk.YES)
        self.results_treeview.column("Puntos", width=100, anchor=tk.CENTER, stretch=tk.NO)

        # Scrollbars for Treeview
        vsb = ttk.Scrollbar(results_frame, orient="vertical", command=self.results_treeview.yview)
        hsb = ttk.Scrollbar(results_frame, orient="horizontal", command=self.results_treeview.xview)
        self.results_treeview.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        # Use grid for treeview and scrollbars to manage space better
        results_frame.grid_rowconfigure(0, weight=1)
        results_frame.grid_columnconfigure(0, weight=1)

        self.results_treeview.grid(row=0, column=0, sticky="nsew", padx=(0,5), pady=(0,5)) # Pad to not touch scrollbars
        vsb.grid(row=0, column=1, sticky="ns", pady=(0,5))
        hsb.grid(row=1, column=0, sticky="ew", padx=(0,5))


        # DOCX Generation Frame
        docx_frame = ttk.LabelFrame(main_frame, text="3. DOCX Generation", padding="10 10 10 10")
        docx_frame.pack(fill=tk.X, pady=10) # Increased pady

        # Template Selection
        select_template_button = ttk.Button(docx_frame, text="Select DOCX Template", command=self.select_docx_template)
        select_template_button.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.docx_template_path_label = ttk.Label(docx_frame, text="No DOCX template selected")
        self.docx_template_path_label.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        # Frame for dynamically generated fields from DOCX template
        self.dynamic_fields_frame = ttk.LabelFrame(docx_frame, text="Template Placeholders", padding="10 10 10 10") # Added LabelFrame
        self.dynamic_fields_frame.grid(row=1, column=0, columnspan=2, sticky="ew", pady=10, padx=5) # Added padx

        # Generate Document Button
        self.generate_doc_button = ttk.Button(docx_frame, text="Generate Document from Template", command=self.generate_document)
        self.generate_doc_button.grid(row=2, column=0, columnspan=2, pady=(10,5)) # pady top and bottom
        self.generate_doc_button.config(state=tk.DISABLED)

        docx_frame.grid_columnconfigure(1, weight=1) # Ensure label/entry expands


        # Status Frame / Bar
        status_frame = ttk.LabelFrame(main_frame, text="Application Status", padding="10 10 10 10") # Increased padding
        status_frame.pack(fill=tk.X, pady=(10,0)) # pady top, no bottom
        self.status_label = ttk.Label(status_frame, textvariable=self.status_var, wraplength=780) # Allow wrapping
        self.status_label.pack(fill=tk.X, padx=5, pady=5)
        self.status_var.set("Ready. Select PDF and enter Subject Code to start.")


    def select_pdf_file(self):
        filepath = filedialog.askopenfilename(
            title="Select PDF File",
            filetypes=(("PDF files", "*.pdf"), ("All files", "*.*"))
        )
        if filepath:
            self.pdf_path.set(filepath)
            self.pdf_path_label.config(text=os.path.basename(filepath))
            self.status_var.set(f"Selected PDF: {os.path.basename(filepath)}")
        else:
            self.pdf_path.set("")
            self.pdf_path_label.config(text="No PDF selected")
            self.status_var.set("PDF selection cancelled.")
        self.check_enable_generate_button()

    def process_extraction(self):
        self.status_var.set("Processing...")
        self.current_json_path = None # Reset
        self.check_enable_generate_button() # Update button state

        # Clear previous results from Treeview
        for item in self.results_treeview.get_children():
            self.results_treeview.delete(item)

        pdf_path_str = self.pdf_path.get()
        subject_code_str = self.subject_code_var.get().strip()

        if not pdf_path_str or not os.path.exists(pdf_path_str):
            self.status_var.set("Error: Please select a valid PDF file.")
            messagebox.showerror("Input Error", "Please select a valid PDF file.")
            return

        if not subject_code_str:
            self.status_var.set("Error: Subject Code is required.")
            messagebox.showerror("Input Error", "Subject Code is required.")
            return

        try:
            self.status_var.set(f"Extracting data for code '{subject_code_str}' from '{os.path.basename(pdf_path_str)}'...")
            self.root.update_idletasks()

            nombre_materia, registros = extraer_por_codigo(pdf_path_str, subject_code_str)

            if nombre_materia is None or not nombre_materia.strip():
                self.status_var.set(f"Subject code '{subject_code_str}' not found in the PDF.")
                messagebox.showinfo("Extraction Info", f"Subject code '{subject_code_str}' not found in the PDF.")
                return

            if not registros:
                self.status_var.set(f"No student records found for subject '{nombre_materia}' (Code: {subject_code_str}).")
                messagebox.showinfo("Extraction Info", f"No student records found for subject '{nombre_materia}' (Code: {subject_code_str}).")
                return

            for reg in registros:
                self.results_treeview.insert("", tk.END, values=(reg.get("Orden"), reg.get("Apellido y Nombre"), reg.get("Puntos")))

            self.status_var.set(f"Found {len(registros)} records for '{nombre_materia}'. Saving data...")
            self.root.update_idletasks()

            output_directory = "output_data" # Changed output directory
            os.makedirs(output_directory, exist_ok=True) # Ensure directory exists
            csv_path, json_path = save_extracted_data(subject_code_str, nombre_materia, registros, output_dir=output_directory)
            self.current_json_path = json_path

            self.status_var.set(f"{len(registros)} records extracted. Saved to '{output_directory}' folder. JSON: {os.path.basename(json_path)}")
            messagebox.showinfo("Extraction Successful", f"{len(registros)} records extracted for '{nombre_materia}'.\nFiles saved in '{output_directory}'.\nCSV: {os.path.basename(csv_path)}\nJSON: {os.path.basename(json_path)}")

        except Exception as e:
            self.status_var.set(f"Error during PDF extraction: {e}") # Clarified source of error
            messagebox.showerror("PDF Extraction Error", f"An error occurred during data extraction from PDF:\n{e}")
        finally:
            self.check_enable_generate_button()

    def select_docx_template(self):
        filepath = filedialog.askopenfilename(
            title="Select DOCX Template File",
            filetypes=(("Word Documents", "*.docx"), ("All files", "*.*"))
        )
        if filepath:
            self.docx_template_path.set(filepath)
            self.docx_template_path_label.config(text=os.path.basename(filepath))
            self.status_var.set(f"Selected DOCX Template: {os.path.basename(filepath)}")
            self.populate_dynamic_fields(filepath)
        else:
            self.docx_template_path.set("")
            self.docx_template_path_label.config(text="No DOCX template selected")
            for widget in self.dynamic_fields_frame.winfo_children(): # Clear dynamic fields
                widget.destroy()
            self.docx_dynamic_fields_entries.clear()
            self.status_var.set("DOCX template selection cancelled.")
        self.check_enable_generate_button()

    def populate_dynamic_fields(self, template_path):
        for widget in self.dynamic_fields_frame.winfo_children(): # Clear previous
            widget.destroy()
        self.docx_dynamic_fields_entries.clear()

        try:
            doc = Document(template_path) # This might raise if file is invalid docx
            campos = detectar_campos(doc)

            if not campos:
                no_fields_label = ttk.Label(self.dynamic_fields_frame, text="No custom fields (e.g., _FIELDNAME) found in template.")
                no_fields_label.pack(pady=10, padx=5) # Added padding
                return

            # Bold title for this section
            title_label = ttk.Label(self.dynamic_fields_frame, text="Fill in the template fields:", font="-weight bold")
            title_label.pack(pady=(5,10), padx=5, anchor="w") # pady top/bottom

            for idx, field_name in enumerate(sorted(list(campos))):
                if field_name.upper() == "_MATERIA":
                    continue

                field_row_frame = ttk.Frame(self.dynamic_fields_frame) # Frame for each row for better layout control
                field_row_frame.pack(fill=tk.X, pady=3, padx=5)

                display_name = field_name.replace("_", " ").strip().title()

                label = ttk.Label(field_row_frame, text=f"{display_name}:", width=25) # Fixed width for alignment
                label.pack(side=tk.LEFT, padx=(0,5)) # Pad only to the right of label

                entry_var = tk.StringVar()
                entry = ttk.Entry(field_row_frame, textvariable=entry_var) # Removed fixed width, let it expand
                entry.pack(side=tk.LEFT, expand=True, fill=tk.X)

                self.docx_dynamic_fields_entries[field_name] = entry_var

        except Exception as e: # Catches errors from Document() or other issues
            self.status_var.set(f"Error processing DOCX template: {e}")
            messagebox.showerror("Template Error", f"Error processing DOCX template. Ensure it's a valid .docx file.\nDetail: {e}")
            for widget in self.dynamic_fields_frame.winfo_children():
                widget.destroy()
            self.docx_dynamic_fields_entries.clear()

    def generate_document(self):
        # Validations
        if not self.current_json_path or not os.path.exists(self.current_json_path):
            self.status_var.set("Error: Extracted data (JSON) not found. Please extract data first.")
            messagebox.showerror("Data Missing", "Extracted data (JSON) not found. Please run 'Extract Data' first.")
            return

        template_path_str = self.docx_template_path.get()
        if not template_path_str or not os.path.exists(template_path_str):
            self.status_var.set("Error: DOCX template not selected or not found.")
            messagebox.showerror("Template Missing", "DOCX template not selected or not found. Please select a template.")
            return

        self.status_var.set("Generating document...")
        self.root.update_idletasks()

        try:
            # This call is now wrapped and can raise FileNotFoundError, ValueError, IOError
            json_data = cargar_json(self.current_json_path)
            # No need to check `if not json_data` as exceptions are raised by cargar_json

            user_provided_data = {}
            for field_name, entry_var in self.docx_dynamic_fields_entries.items():
                user_provided_data[field_name] = entry_var.get()

            final_datos_for_docx = {
                "_MATERIA": json_data.get("materia", "MATERIA NO ESPECIFICADA")
            }
            final_datos_for_docx.update(user_provided_data)

            temp_doc_check = Document(template_path_str) # Use validated path
            all_template_fields = detectar_campos(temp_doc_check)
            for field in all_template_fields:
                if field not in final_datos_for_docx:
                    final_datos_for_docx[field] = ""

            doc = Document(template_path_str) # Use validated path
            reemplazar_campos(doc, final_datos_for_docx)
            completar_listado(doc, json_data.get("listado", []))

            output_docs_dir = "output_documents" # New output directory for DOCX
            os.makedirs(output_docs_dir, exist_ok=True) # Ensure directory exists

            base_name = os.path.basename(self.current_json_path)
            name_part = base_name.replace("_listado.json", "").replace("_resumido.csv","") # Handle if json name changes
            output_filename = f"{name_part}_designacion.docx"
            output_docx_path = os.path.join(output_docs_dir, output_filename)

            doc.save(output_docx_path)
            self.status_var.set(f"Document generated: {output_docx_path}")
            messagebox.showinfo("Success", f"Document generated successfully!\nSaved in '{output_docs_dir}' folder as: {output_filename}")

        except FileNotFoundError as e: # Specific error from cargar_json or Document()
            self.status_var.set(f"Error: File not found. {e}")
            messagebox.showerror("File Not Found Error", f"A required file was not found.\n{e}")
        except ValueError as e: # Specific error from cargar_json (JSONDecodeError)
            self.status_var.set(f"Error: Invalid data format. {e}")
            messagebox.showerror("Data Format Error", f"There was an issue with the data format.\n{e}")
        except IOError as e: # Specific error from cargar_json or file operations
            self.status_var.set(f"Error: File operation failed. {e}")
            messagebox.showerror("File Operation Error", f"A file operation failed.\n{e}")
        except Exception as e: # Catch other errors during doc generation or template processing
            self.status_var.set(f"Error generating document: {e}")
            messagebox.showerror("Generation Error", f"An error occurred while generating the document:\n{e}")
        finally:
            self.check_enable_generate_button()

    def check_enable_generate_button(self):
        if self.current_json_path and \
           os.path.exists(self.current_json_path) and \
           self.docx_template_path.get() and \
           os.path.exists(self.docx_template_path.get()):
            self.generate_doc_button.config(state=tk.NORMAL)
        else:
            self.generate_doc_button.config(state=tk.DISABLED)


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
