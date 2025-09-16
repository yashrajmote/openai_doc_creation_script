#!/usr/bin/env python3
"""
Drag and Drop Desktop Application for Excel to DOCX Generator
A simple drag-and-drop interface using tkinter with file drop functionality.
"""

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import tkinterdnd2 as tkdnd
import os
import threading
from excel_to_docx_generator import ExcelToDocxGenerator


class DragDropApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel to DOCX Generator - Drag & Drop")
        self.root.geometry("500x400")
        self.root.resizable(True, True)
        
        # Variables
        self.output_directory = tk.StringVar()
        self.processing = False
        
        self.setup_ui()
        
    def setup_ui(self):
        """Set up the user interface."""
        # Main frame
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        title_label = ttk.Label(main_frame, text="📄 Excel to DOCX Generator", 
                               font=("Arial", 18, "bold"))
        title_label.pack(pady=(0, 20))
        
        # Instructions
        instructions = ttk.Label(main_frame, 
                               text="Drag and drop your Excel file here\nor click to browse",
                               font=("Arial", 12),
                               foreground="gray")
        instructions.pack(pady=10)
        
        # Process button - MOVED ABOVE DROP ZONE
        self.process_button = ttk.Button(main_frame, text="🚀 Generate DOCX Files", 
                                       command=self.process_files, 
                                       state="disabled",
                                       style="Accent.TButton")
        self.process_button.pack(pady=15)
        
        # Output directory
        dir_frame = ttk.Frame(main_frame)
        dir_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(dir_frame, text="Output Directory:").pack(side=tk.LEFT)
        ttk.Entry(dir_frame, textvariable=self.output_directory, width=30).pack(side=tk.LEFT, padx=(5, 5))
        ttk.Button(dir_frame, text="Browse", command=self.browse_output_directory).pack(side=tk.LEFT)
        
        # Drop zone
        self.drop_zone = tk.Frame(main_frame, 
                                 bg="lightgray", 
                                 relief="ridge", 
                                 bd=2,
                                 height=120)
        self.drop_zone.pack(fill=tk.BOTH, expand=True, pady=10)
        self.drop_zone.pack_propagate(False)
        
        # Drop zone label
        self.drop_label = ttk.Label(self.drop_zone, 
                                   text="📁 Drop Excel file here",
                                   font=("Arial", 14),
                                   background="lightgray")
        self.drop_label.pack(expand=True)
        
        # File browser button inside drop zone
        browse_button = ttk.Button(self.drop_zone, 
                                  text="📂 Or click to browse",
                                  command=self.browse_excel_file,
                                  style="Accent.TButton")
        browse_button.pack(pady=10)
        
        # Configure drag and drop
        self.drop_zone.drop_target_register(tkdnd.DND_FILES)
        self.drop_zone.dnd_bind('<<Drop>>', self.on_drop)
        
        # Status
        self.status_label = ttk.Label(main_frame, text="Ready - Drop an Excel file to begin")
        self.status_label.pack()
        
        # Results area
        self.results_text = scrolledtext.ScrolledText(main_frame, height=8, width=60)
        self.results_text.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # Set default output directory
        self.output_directory.set("/Users/yash/Desktop/Desktop - YASH's MacBook Air/APPLICATIONS")
        
        # Store current file
        self.current_file = None
        
    def on_drop(self, event):
        """Handle file drop event."""
        files = self.root.tk.splitlist(event.data)
        if files:
            file_path = files[0]
            if file_path.lower().endswith(('.xlsx', '.xls')):
                self.current_file = file_path
                self.drop_label.config(text=f"✅ {os.path.basename(file_path)}")
                self.process_button.config(state="normal")
                self.status_label.config(text="File ready - Click 'Generate DOCX Files' to process")
            else:
                messagebox.showerror("Error", "Please drop an Excel file (.xlsx or .xls)")
                
    def browse_excel_file(self):
        """Browse for Excel file."""
        from tkinter import filedialog
        filename = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            self.current_file = filename
            self.drop_label.config(text=f"✅ {os.path.basename(filename)}")
            self.process_button.config(state="normal")
            self.status_label.config(text="File ready - Click 'Generate DOCX Files' to process")
            
    def browse_output_directory(self):
        """Browse for output directory."""
        from tkinter import filedialog
        directory = filedialog.askdirectory(title="Select Output Directory")
        if directory:
            self.output_directory.set(directory)
            
    def process_files(self):
        """Process the Excel file and generate DOCX files."""
        if not self.current_file:
            messagebox.showerror("Error", "Please drop an Excel file first.")
            return
            
        if not self.output_directory.get():
            messagebox.showerror("Error", "Please select an output directory.")
            return
            
        if not os.path.exists(self.current_file):
            messagebox.showerror("Error", "Excel file does not exist.")
            return
            
        # Start processing in a separate thread
        self.processing = True
        self.process_button.config(state="disabled")
        self.status_label.config(text="Processing...")
        self.results_text.delete(1.0, tk.END)
        
        # Run processing in separate thread
        thread = threading.Thread(target=self.run_processing)
        thread.daemon = True
        thread.start()
        
    def run_processing(self):
        """Run the actual processing in a separate thread."""
        try:
            # Create generator and process file
            generator = ExcelToDocxGenerator(
                self.current_file, 
                self.output_directory.get()
            )
            
            results = generator.process_excel_file()
            
            if results:
                # Update UI with results
                self.root.after(0, self.update_results, results)
            else:
                self.root.after(0, self.show_error, "Processing failed. Please check the error messages.")
                
        except Exception as e:
            self.root.after(0, self.show_error, f"Error: {str(e)}")
        finally:
            self.root.after(0, self.processing_complete)
            
    def update_results(self, results):
        """Update the UI with processing results."""
        self.results_text.delete(1.0, tk.END)
        
        # Display results
        self.results_text.insert(tk.END, "🎉 PROCESSING COMPLETE!\n")
        self.results_text.insert(tk.END, "=" * 50 + "\n")
        self.results_text.insert(tk.END, f"📊 Total rows processed: {results['total_rows']}\n")
        self.results_text.insert(tk.END, f"✅ Valid entries: {results['valid_entries']}\n")
        self.results_text.insert(tk.END, f"⏭️ Skipped entries: {results['skipped_entries']}\n")
        self.results_text.insert(tk.END, f"📄 DOCX files created: {len(results['created_files'])}\n")
        self.results_text.insert(tk.END, f"📁 Output directory: {self.output_directory.get()}\n\n")
        
        if results['created_files']:
            self.results_text.insert(tk.END, "📋 Created files:\n")
            for filepath in results['created_files']:
                filename = os.path.basename(filepath)
                self.results_text.insert(tk.END, f"  • {filename}\n")
        
        self.status_label.config(text=f"✅ Success! Created {len(results['created_files'])} DOCX files")
        
    def show_error(self, message):
        """Show error message."""
        self.results_text.delete(1.0, tk.END)
        self.results_text.insert(tk.END, f"❌ ERROR: {message}")
        self.status_label.config(text="❌ Processing failed")
        
    def processing_complete(self):
        """Called when processing is complete."""
        self.processing = False
        self.process_button.config(state="normal")


def main():
    """Main function to run the drag and drop application."""
    root = tkdnd.Tk()
    
    # Configure style
    style = ttk.Style()
    style.theme_use('clam')
    
    # Create and run the application
    app = DragDropApp(root)
    
    # Center the window
    root.update_idletasks()
    x = (root.winfo_screenwidth() // 2) - (root.winfo_width() // 2)
    y = (root.winfo_screenheight() // 2) - (root.winfo_height() // 2)
    root.geometry(f"+{x}+{y}")
    
    root.mainloop()


if __name__ == "__main__":
    main()
