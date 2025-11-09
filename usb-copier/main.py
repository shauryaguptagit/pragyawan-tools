import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import subprocess
import threading
from datetime import datetime

try:
    import psutil
except ImportError:
    messagebox.showerror(
        "Dependency Missing",
        "The 'psutil' library is required. Please install it by running:\npip install psutil"
    )
    exit()

class USBCopierApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Ironclad USB Drive Copier - created by Shaurya Gupta")
        self.root.geometry("700x650")
        self.root.minsize(600, 500)
        
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        # --- Variables ---
        self.source_files = []
        self.drive_vars = {}
        self.copy_in_progress = False
        self.verify_copy = tk.BooleanVar(value=False) # Default to OFF for speed

        self.create_widgets()
        self.scan_drives()
        
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(4, weight=1) # Let the log area expand

        title_label = ttk.Label(main_frame, text="Ironclad USB Drive Copier - Shaurya Gupta", font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20), sticky=tk.W)

        files_frame = ttk.LabelFrame(main_frame, text="Source Files")
        files_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        files_frame.columnconfigure(0, weight=1)
        
        self.file_listbox = tk.Listbox(files_frame, height=5, selectmode=tk.EXTENDED)
        self.file_listbox.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5, pady=5)
        
        file_scrollbar = ttk.Scrollbar(files_frame, orient=tk.VERTICAL, command=self.file_listbox.yview)
        file_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S), pady=5)
        self.file_listbox.config(yscrollcommand=file_scrollbar.set)
        
        file_buttons_frame = ttk.Frame(files_frame)
        file_buttons_frame.grid(row=0, column=2, sticky=(tk.N, tk.S), padx=5)
        
        self.add_files_button = ttk.Button(file_buttons_frame, text="Add Files", command=self.add_files)
        self.add_files_button.pack(fill=tk.X, pady=2)
        self.remove_files_button = ttk.Button(file_buttons_frame, text="Remove", command=self.remove_selected_files)
        self.remove_files_button.pack(fill=tk.X, pady=2)

        self.drives_frame = ttk.LabelFrame(main_frame, text="Available Removable Drives")
        self.drives_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10)
        self.drives_frame.columnconfigure(1, weight=1)

        progress_frame = ttk.LabelFrame(main_frame, text="Progress")
        progress_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        progress_frame.columnconfigure(0, weight=1)
        
        self.progress_bar = ttk.Progressbar(progress_frame, mode='determinate')
        self.progress_bar.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=10, pady=5)
        
        self.status_label = ttk.Label(progress_frame, text="Ready")
        self.status_label.grid(row=1, column=0, sticky=tk.W, padx=10, pady=(0, 5))

        log_frame = ttk.LabelFrame(main_frame, text="Log")
        log_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        self.log_text = tk.Text(log_frame, height=8, wrap=tk.WORD, relief=tk.FLAT)
        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=5, column=0, columnspan=2, pady=10, sticky=(tk.E, tk.W))
        button_frame.columnconfigure(1, weight=1)

        self.verify_checkbox = ttk.Checkbutton(
            button_frame, 
            text="Verify files after copying (slower but safer)",
            variable=self.verify_copy
        )
        self.verify_checkbox.grid(row=0, column=0, sticky=tk.W, padx=5)

        action_buttons_frame = ttk.Frame(button_frame)
        action_buttons_frame.grid(row=0, column=1, sticky=tk.E)
        
        self.start_button = ttk.Button(action_buttons_frame, text="Start Copying", command=self.start_copy, style="Accent.TButton")
        self.start_button.pack(side=tk.LEFT, padx=5)
        self.clear_log_button = ttk.Button(action_buttons_frame, text="Clear Log", command=self.clear_log)
        self.clear_log_button.pack(side=tk.LEFT, padx=5)
        ttk.Button(action_buttons_frame, text="Exit", command=self.on_closing).pack(side=tk.LEFT, padx=5)
        self.style.configure("Accent.TButton", font=("Arial", 10, "bold"))
        
        # List of controls to disable during copy
        self.controls_to_disable = [
            self.add_files_button, self.remove_files_button, self.verify_checkbox, 
            self.start_button
        ]

    def add_files(self):
        filenames = filedialog.askopenfilenames(title="Select files to copy")
        if filenames:
            for f in filenames:
                if f not in self.source_files:
                    self.source_files.append(f)
                    self.file_listbox.insert(tk.END, os.path.basename(f))
                    self.log_message(f"Added file: {os.path.basename(f)}")

    def remove_selected_files(self):
        selected_indices = self.file_listbox.curselection()
        # Iterate backwards to avoid index shifting issues when deleting
        for i in sorted(selected_indices, reverse=True):
            self.log_message(f"Removed file: {os.path.basename(self.source_files[i])}")
            self.file_listbox.delete(i)
            del self.source_files[i]

    def scan_drives(self):
        self.log_message("Scanning for removable drives...")
        for widget in self.drives_frame.winfo_children():
            widget.destroy()
        self.drive_vars.clear()
        partitions = psutil.disk_partitions()
        removable_drives = [p for p in partitions if 'removable' in p.opts or 'cdrom' in p.opts]
        if not removable_drives:
            ttk.Label(self.drives_frame, text="No removable drives found.").grid(row=0, column=0, padx=10, pady=10)
        else:
            drive_selection_frame = ttk.Frame(self.drives_frame)
            drive_selection_frame.grid(row=0, column=0, sticky=tk.W, padx=10, pady=5)
            row, col = 0, 0
            for p in removable_drives:
                drive_letter = p.device.replace('\\', '')
                self.drive_vars[drive_letter] = tk.BooleanVar(value=True)
                chk = ttk.Checkbutton(drive_selection_frame, text=f"{drive_letter} ({p.fstype})", variable=self.drive_vars[drive_letter])
                chk.grid(row=row, column=col, sticky=tk.W, padx=5, pady=2)
                col += 1
                if col % 4 == 0: col = 0; row += 1
            drive_buttons_frame = ttk.Frame(self.drives_frame)
            drive_buttons_frame.grid(row=1, column=0, sticky=tk.W, padx=10, pady=5)
            self.select_all_button = ttk.Button(drive_buttons_frame, text="Select All", command=lambda: self.toggle_all_drives(True))
            self.select_all_button.pack(side=tk.LEFT)
            self.deselect_all_button = ttk.Button(drive_buttons_frame, text="Deselect All", command=lambda: self.toggle_all_drives(False))
            self.deselect_all_button.pack(side=tk.LEFT, padx=5)
            self.refresh_drives_button = ttk.Button(self.drives_frame, text="Refresh Drives", command=self.scan_drives)
            self.refresh_drives_button.grid(row=0, column=1, rowspan=2, padx=10, sticky=tk.E)
            self.controls_to_disable.extend([self.select_all_button, self.deselect_all_button, self.refresh_drives_button])

    def toggle_controls(self, enable):
        state = tk.NORMAL if enable else tk.DISABLED
        for control in self.controls_to_disable:
            control.config(state=state)
        # Also disable drive checkboxes
        for child in self.drives_frame.winfo_children():
            try:
                child.config(state=state)
                for grandchild in child.winfo_children():
                    grandchild.config(state=state)
            except tk.TclError:
                pass

    def start_copy(self):
        if not self.source_files: 
            messagebox.showerror("Error", "Please add at least one source file.")
            return
        selected_drives = [drive for drive, var in self.drive_vars.items() if var.get()]
        if not selected_drives: 
            messagebox.showerror("Error", "Please select at least one drive.")
            return

        confirm = messagebox.askyesno("Confirm Copy", f"Copy {len(self.source_files)} file(s) to {len(selected_drives)} drive(s)?")
        
        if confirm:
            self.copy_in_progress = True
            self.toggle_controls(enable=False)
            
            total_operations = len(self.source_files) * len(selected_drives)
            self.progress_bar['maximum'] = total_operations
            self.progress_bar['value'] = 0
            
            self.log_message(f"Starting copy of {total_operations} total file operations.")
            if self.verify_copy.get():
                self.log_message("NOTE: File verification is ON. This will be slower but safer.")

            thread = threading.Thread(
                target=self.copy_files_thread, 
                args=(self.source_files.copy(), selected_drives, self.verify_copy.get())
            )
            thread.daemon = True
            thread.start()

    def copy_files_thread(self, source_files, drives, verify):
        successful_ops = 0
        for i, drive in enumerate(drives):
            self.root.after(0, self.update_status, f"Preparing to copy to {drive}...")
            target_drive_path = f"{drive}\\"
            
            for j, file_path in enumerate(source_files):
                filename = os.path.basename(file_path)
                source_dir = os.path.dirname(file_path)
                
                progress_text = f"Copying '{filename}' to {drive} ({j+1}/{len(source_files)})"
                self.root.after(0, self.update_status, progress_text)
                
                try:
                    cmd = ['robocopy', source_dir, target_drive_path, filename, '/R:1', '/W:2']
                    if verify:
                        cmd.append('/V') # Add verification flag if requested

                    result = subprocess.run(cmd, capture_output=True, text=True, check=False, creationflags=subprocess.CREATE_NO_WINDOW)
                    
                    if result.returncode < 8:
                        successful_ops += 1
                    else:
                        self.log_message(f"✗ FAILED to copy '{filename}' to {drive} (Code: {result.returncode})")
                        self.log_message(f"  Details: {result.stdout.strip()} {result.stderr.strip()}")
                
                except Exception as e:
                    self.log_message(f"✗ CRITICAL ERROR copying '{filename}' to {drive}: {e}")
                
                self.root.after(0, self.progress_bar.step)
        
        total_ops = len(source_files) * len(drives)
        self.root.after(0, self.copy_complete, successful_ops, total_ops)

    def update_status(self, text):
        self.status_label.config(text=text)

    def copy_complete(self, successful_ops, total_ops):
        self.copy_in_progress = False
        self.toggle_controls(enable=True)
        self.status_label.config(text="Completed")
        
        message = f"Operation complete. Successfully copied {successful_ops} of {total_ops} files."
        self.log_message(message)
        
        if successful_ops < total_ops:
            messagebox.showwarning("Completed with Errors", message)
        else:
            messagebox.showinfo("Completed Successfully", message)

    def log_message(self, message):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()
        
    def clear_log(self):
        self.log_text.delete(1.0, tk.END)
        self.log_message("Log cleared.")
        
    def on_closing(self):
        if self.copy_in_progress:
            if messagebox.askyesno("Confirm Exit", "A copy operation is in progress. Are you sure you want to exit?"): 
                self.root.destroy()
        else:
            self.root.destroy()
            
    def toggle_all_drives(self, select):
        for var in self.drive_vars.values(): 
            var.set(select)

def main():
    root = tk.Tk()
    app = USBCopierApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()