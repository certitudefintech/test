import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import difflib
import customtkinter as ctk
import threading
import time
import os
from datetime import datetime
import calendar

class LoadingWindow(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Processing")
        
        # Configure window
        self.geometry("300x150")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        
        # Center the window
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f'{width}x{height}+{x}+{y}')
        
        # Create loading frame
        self.loading_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.loading_frame.pack(expand=True, fill="both", padx=20, pady=20)
        
        # Loading label
        self.loading_label = ctk.CTkLabel(
            self.loading_frame,
            text="Processing your files...",
            font=("Segoe UI", 14, "bold"),
            text_color="#ffffff"
        )
        self.loading_label.pack(pady=(0, 15))
        
        # Progress bar
        self.progress_bar = ctk.CTkProgressBar(self.loading_frame)
        self.progress_bar.pack(fill="x", padx=20)
        self.progress_bar.set(0)
        
        # Status label
        self.status_label = ctk.CTkLabel(
            self.loading_frame,
            text="Initializing...",
            font=("Segoe UI", 12),
            text_color="#b0b0b0"
        )
        self.status_label.pack(pady=(10, 0))
        
        # Start animation
        self.animation_running = True
        self.animate()

    def animate(self):
        if not self.animation_running:
            return
        
        # Update progress bar
        current = self.progress_bar.get()
        if current >= 1:
            self.progress_bar.set(0)
        else:
            self.progress_bar.set(current + 0.1)
        
        # Schedule next update
        self.after(100, self.animate)

    def update_status(self, status):
        self.status_label.configure(text=status)

    def stop(self):
        self.animation_running = False
        self.destroy()

class SwitchRegisterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Switch Register - Launcher")
        self.root.geometry("800x700")
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
        self.root.resizable(False, False)
        
        # Store selected file paths
        self.switch_register_path = None
        self.rta_master_path = None
        self.brokerage_structure_paths = []  # List to store multiple files
        
        # Create UI
        self.create_widgets()
        
        # Center the window
        self.center_window()
    
    def center_window(self):
        """Center the window on screen"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
    
    def create_widgets(self):
        """Create and arrange all GUI widgets"""
        # Main container with dark background
        main_frame = ctk.CTkFrame(self.root, fg_color="#1a1a1a")
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Title section
        title_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        title_frame.pack(pady=(20, 10))
        
        title_label = ctk.CTkLabel(
            title_frame,
            text="SWITCH REGISTER",
            font=("Segoe UI", 32, "bold"),
            text_color="#ffffff"
        )
        title_label.pack()
        
        subtitle_label = ctk.CTkLabel(
            title_frame,
            text="Upload Center",
            font=("Segoe UI", 16),
            text_color="#b0b0b0"
        )
        subtitle_label.pack(pady=(5, 0))
        
        # Central panel with rounded corners
        upload_panel = ctk.CTkFrame(main_frame, fg_color="#2b2b2b", corner_radius=15)
        upload_panel.pack(fill="both", expand=True, padx=30, pady=20)
        
        # Instruction text
        instruction_label = ctk.CTkLabel(
            upload_panel,
            text="Select files to begin",
            font=("Segoe UI", 14),
            text_color="#b0b0b0"
        )
        instruction_label.pack(pady=(30, 30))
        
        # File upload sections
        # 1. Switch Register (Blue)
        self.create_upload_section(
            upload_panel,
            "SWITCH REGISTER",
            "#3498db",
            "#2980b9",
            self.upload_switch_register,
            "switch_register"
        )
        
        # 2. RTA Master (Light Blue/Cyan)
        self.create_upload_section(
            upload_panel,
            "RTA MASTER",
            "#5dade2",
            "#3498db",
            self.upload_rta_master,
            "rta_master"
        )
        
        # 3. Brokerage Structure File (Green) - Multiple files
        self.create_upload_section(
            upload_panel,
            "BROKERAGE STRUCTURE FILE",
            "#27ae60",
            "#219a52",
            self.upload_brokerage_structure,
            "brokerage_structure",
            multiple=True
        )
        
        # Process button at bottom
        process_btn = ctk.CTkButton(
            upload_panel,
            text="PROCESS",
            command=self.process_files,
            fg_color="#27ae60",
            hover_color="#219a52",
            font=("Segoe UI", 16, "bold"),
            height=50,
            corner_radius=10
        )
        process_btn.pack(fill="x", padx=30, pady=(20, 30))
        
        # Status bar
        self.status_label = ctk.CTkLabel(
            main_frame,
            text="Ready to upload files",
            font=("Segoe UI", 11),
            text_color="#b0b0b0"
        )
        self.status_label.pack(pady=(0, 10))
    
    def create_upload_section(self, parent, label_text, button_color, hover_color, command, file_type, multiple=False):
        """Create a file upload section with button and status label"""
        section_frame = ctk.CTkFrame(parent, fg_color="transparent")
        section_frame.pack(fill="x", padx=30, pady=10)
        
        # Upload button
        upload_btn = ctk.CTkButton(
            section_frame,
            text=label_text,
            command=command,
            fg_color=button_color,
            hover_color=hover_color,
            font=("Segoe UI", 12, "bold"),
            width=200,
            height=40,
            corner_radius=8
        )
        upload_btn.pack(side="left", padx=(0, 15))
        
        # Status label
        status_label = ctk.CTkLabel(
            section_frame,
            text="No file selected" if not multiple else "No files selected",
            text_color="#b0b0b0",
            font=("Segoe UI", 11),
            anchor="w"
        )
        status_label.pack(side="left", fill="x", expand=True)
        
        # Store reference to status label
        if file_type == "switch_register":
            self.switch_register_status = status_label
        elif file_type == "rta_master":
            self.rta_master_status = status_label
        elif file_type == "brokerage_structure":
            self.brokerage_structure_status = status_label
    
    def upload_switch_register(self):
        """Handle Switch Register file upload"""
        file_path = filedialog.askopenfilename(
            title="Select Switch Register File",
            filetypes=[
                ("Excel files", "*.xlsx *.xls"),
                ("CSV files", "*.csv"),
                ("All files", "*.*")
            ]
        )
        if file_path:
            self.switch_register_path = file_path
            filename = os.path.basename(file_path)
            self.switch_register_status.configure(
                text=f"✓ {filename}",
                text_color="#3498db"
            )
            self.update_status()
            messagebox.showinfo("Success", "Switch Register file uploaded successfully!")
    
    def upload_rta_master(self):
        """Handle RTA Master file upload"""
        file_path = filedialog.askopenfilename(
            title="Select RTA Master File",
            filetypes=[
                ("Excel files", "*.xlsx *.xls"),
                ("CSV files", "*.csv"),
                ("All files", "*.*")
            ]
        )
        if file_path:
            self.rta_master_path = file_path
            filename = os.path.basename(file_path)
            self.rta_master_status.configure(
                text=f"✓ {filename}",
                text_color="#5dade2"
            )
            self.update_status()
            messagebox.showinfo("Success", "RTA Master file uploaded successfully!")
    
    def upload_brokerage_structure(self):
        """Handle Brokerage Structure file upload (multiple files)"""
        file_paths = filedialog.askopenfilenames(
            title="Select Brokerage Structure File(s)",
            filetypes=[
                ("Excel files", "*.xlsx *.xls"),
                ("CSV files", "*.csv"),
                ("All files", "*.*")
            ]
        )
        if file_paths:
            self.brokerage_structure_paths = list(file_paths)
            if len(file_paths) == 1:
                filename = os.path.basename(file_paths[0])
                self.brokerage_structure_status.configure(
                    text=f"✓ {filename}",
                    text_color="#27ae60"
                )
            else:
                self.brokerage_structure_status.configure(
                    text=f"✓ {len(file_paths)} files selected",
                    text_color="#27ae60"
                )
            self.update_status()
            messagebox.showinfo("Success", f"{len(file_paths)} Brokerage Structure file(s) uploaded successfully!")
    
    def update_status(self):
        """Update the main status label"""
        files_uploaded = sum([
            self.switch_register_path is not None,
            self.rta_master_path is not None,
            len(self.brokerage_structure_paths) > 0
        ])
        
        if files_uploaded == 0:
            self.status_label.configure(
                text="Ready to upload files",
                text_color="#b0b0b0"
            )
        elif files_uploaded < 3:
            self.status_label.configure(
                text=f"{files_uploaded} of 3 file types uploaded",
                text_color="#f39c12"
            )
        else:
            self.status_label.configure(
                text="All files uploaded. Ready to process!",
                text_color="#27ae60"
            )
    
    def process_files(self):
        """Handle the process button click"""
        # Check if required files are uploaded
        if not self.switch_register_path:
            messagebox.showwarning(
                "Incomplete Upload",
                "Please upload Switch Register file before processing."
            )
            return
        
        if len(self.brokerage_structure_paths) == 0:
            messagebox.showwarning(
                "Incomplete Upload",
                "Please upload at least one Brokerage Structure file before processing."
            )
            return
        
        # Show loading window
        loading_window = LoadingWindow(self.root)
        
        def process_file():
            try:
                loading_window.update_status("Reading Switch Register file...")
                # Read Switch Register file
                if self.switch_register_path.endswith('.csv'):
                    switch_df = pd.read_csv(self.switch_register_path)
                else:
                    switch_df = pd.read_excel(self.switch_register_path)
                
                # Check for duplicate columns
                if switch_df.columns.duplicated().any():
                    dupes = switch_df.columns[switch_df.columns.duplicated()].tolist()
                    self.root.after(0, lambda: messagebox.showerror(
                        "Error",
                        f"Switch Register file has duplicate column names: {dupes}. Please fix the file and try again."
                    ))
                    self.root.after(0, loading_window.stop)
                    return
                
                # Detect if Switch Register has IN/OUT broker, subfund, and asset columns (no RTA needed)
                # Normalize: strip and collapse spaces so "OUT SUBFUN" / "OUT_SUBFUN" both match
                def _norm(s):
                    try:
                        t = str(s).upper().lstrip('\ufeff').strip()
                        return ' '.join(t.split()).replace(' ', '_')
                    except (TypeError, AttributeError):
                        return ''
                in_broker_col = None
                in_subfund_col = None
                out_broker_col = None
                out_subfund_col = None
                so_asset_col = None
                si_asset_col = None
                for col in switch_df.columns:
                    try:
                        c = _norm(col)
                        c_spaces = c.replace('_', ' ')
                        if c in ('IN_BROKER',) or c_spaces == 'IN BROKER':
                            in_broker_col = col
                        elif c in ('IN_SUBFUND', 'IN_SUBFUN',) or c_spaces in ('IN SUBFUND', 'IN SUBFUN'):
                            in_subfund_col = col
                        elif c in ('OUT_BROKER',) or c_spaces == 'OUT BROKER':
                            out_broker_col = col
                        elif c in ('OUT_SUBFUND', 'OUT_SUBFUN',) or c_spaces in ('OUT SUBFUND', 'OUT SUBFUN'):
                            out_subfund_col = col
                        elif c in ('SO_ASSET_C',) or c_spaces == 'SO ASSET C':
                            so_asset_col = col
                        elif c in ('SI_ASSET_C', 'IN_ASSET_C',) or c_spaces in ('SI ASSET C', 'IN ASSET C'):
                            si_asset_col = col
                    except (TypeError, AttributeError):
                        continue
                use_switch_columns = bool(in_broker_col and in_subfund_col and out_broker_col and out_subfund_col)
                if use_switch_columns:
                    loading_window.update_status("Using broker/subfund/asset from Switch Register (no RTA)...")
                else:
                    if not self.rta_master_path:
                        self.root.after(0, lambda: messagebox.showwarning(
                            "Incomplete Upload",
                            "Please upload RTA Master file, or use a Switch Register with columns: "
                            "IN_BROKER, IN_SUBFUND, OUT_BROKER, OUT_SUBFUN (or OUT_SUBFUND), SO_ASSET_C, SI_ASSET_C (or IN_ASSET_C)."
                        ))
                        self.root.after(0, loading_window.stop)
                        return
                    loading_window.update_status("Reading RTA Master file...")
                    if self.rta_master_path.endswith('.csv'):
                        rta_df = pd.read_csv(self.rta_master_path)
                    else:
                        rta_df = pd.read_excel(self.rta_master_path)
                    if rta_df.columns.duplicated().any():
                        dupes = rta_df.columns[rta_df.columns.duplicated()].tolist()
                        self.root.after(0, lambda: messagebox.showerror(
                            "Error",
                            f"RTA Master file has duplicate column names: {dupes}. Please fix the file and try again."
                        ))
                        self.root.after(0, loading_window.stop)
                        return
                
                loading_window.update_status("Reading Brokerage Structure file(s)...")
                # Read all Brokerage Structure files
                brokerage_dfs = []
                for file_path in self.brokerage_structure_paths:
                    if file_path.endswith('.csv'):
                        df = pd.read_csv(file_path)
                    else:
                        df = pd.read_excel(file_path)
                    
                    # Check for duplicate columns
                    if df.columns.duplicated().any():
                        dupes = df.columns[df.columns.duplicated()].tolist()
                        self.root.after(0, lambda d=dupes, f=os.path.basename(file_path): messagebox.showerror(
                            "Error",
                            f"Brokerage Structure file '{f}' has duplicate column names: {d}. Please fix the file and try again."
                        ))
                        self.root.after(0, loading_window.stop)
                        return
                    
                    brokerage_dfs.append(df)
                
                # Combine all brokerage structure files
                if len(brokerage_dfs) > 1:
                    combined_brokerage_df = pd.concat(brokerage_dfs, ignore_index=True)
                else:
                    combined_brokerage_df = brokerage_dfs[0] if brokerage_dfs else pd.DataFrame()
                
                loading_window.update_status("Processing Switch Register data...")
                
                processed_df = switch_df.copy()
                
                if use_switch_columns:
                    # Use IN/OUT broker, subfund, and asset columns directly from Switch Register (no RTA)
                    processed_df['IN subfund code'] = processed_df[in_subfund_col].astype(str).str.strip()
                    processed_df['out subfund code'] = processed_df[out_subfund_col].astype(str).str.strip()
                    if si_asset_col:
                        processed_df['IN ASSET_CLASS'] = processed_df[si_asset_col].fillna('Not Found').astype(str).str.strip()
                    else:
                        processed_df['IN ASSET_CLASS'] = 'Not Found'
                    if so_asset_col:
                        processed_df['out ASSET_CLASS'] = processed_df[so_asset_col].fillna('Not Found').astype(str).str.strip()
                    else:
                        processed_df['out ASSET_CLASS'] = 'Not Found'
                else:
                    # Process Switch Register: Extract scheme codes from "From" and "Scheme :" columns, then RTA
                    # Function to extract scheme code (part before "/")
                    def extract_scheme_code(value):
                        if pd.isna(value):
                            return None
                        value_str = str(value)
                        if '/' in value_str:
                            scheme_code = value_str.split('/')[0].strip()
                            return scheme_code
                        return None
                    
                    # Find "From" column (case-insensitive)
                    from_col = None
                    for col in processed_df.columns:
                        if str(col).upper().strip() == 'FROM':
                            from_col = col
                            break
                    
                    # Find "Scheme :" column (case-insensitive, handle variations)
                    scheme_col = None
                    for col in processed_df.columns:
                        col_upper = str(col).upper().strip()
                        if col_upper == 'SCHEME :' or col_upper == 'SCHEME:' or col_upper == 'SCHEME':
                            scheme_col = col
                            break
                    
                    # Process "From" column
                    if from_col:
                        # Extract scheme code and add as "out Scheme Code" at the start
                        out_scheme_codes = processed_df[from_col].apply(extract_scheme_code)
                        processed_df.insert(0, 'out Scheme Code', out_scheme_codes)
                        
                        # Rename "From" column to "switch out scheme"
                        processed_df = processed_df.rename(columns={from_col: 'switch out scheme'})
                    else:
                        # If "From" column not found, show warning but continue
                        self.root.after(0, lambda: messagebox.showwarning(
                            "Warning",
                            "From column not found in Switch Register. Processing without out scheme code extraction."
                        ))
                        # Add empty column
                        processed_df.insert(0, 'out Scheme Code', None)
                    
                    # Process "Scheme :" column
                    if scheme_col:
                        # Extract scheme code and add as "IN Scheme Code" after "out Scheme Code"
                        in_scheme_codes = processed_df[scheme_col].apply(extract_scheme_code)
                        # Find position after "out Scheme Code"
                        out_scheme_code_idx = list(processed_df.columns).index('out Scheme Code')
                        processed_df.insert(out_scheme_code_idx + 1, 'IN Scheme Code', in_scheme_codes)
                        
                        # Rename "Scheme :" column to "switch in scheme"
                        processed_df = processed_df.rename(columns={scheme_col: 'switch in scheme'})
                    else:
                        # If "Scheme :" column not found, show warning but continue
                        self.root.after(0, lambda: messagebox.showwarning(
                            "Warning",
                            "Scheme : column not found in Switch Register. Processing without IN scheme code extraction."
                        ))
                        # Add empty column after "out Scheme Code"
                        out_scheme_code_idx = list(processed_df.columns).index('out Scheme Code')
                        processed_df.insert(out_scheme_code_idx + 1, 'IN Scheme Code', None)
                    
                    loading_window.update_status("Matching scheme codes with RTA Master...")
                    
                    # Find Scheme_code, PARENT_SUB_FUND_CODE, and ASSET_CLASS columns in RTA Master
                    scheme_code_col = None
                    parent_sub_fund_code_col = None
                    asset_class_col = None
                    
                    for col in rta_df.columns:
                        try:
                            col_upper = str(col).upper().strip()
                            if col_upper == 'SCHEME_CODE' or col_upper == 'SCHEME CODE':
                                scheme_code_col = col
                            elif col_upper == 'PARENT_SUB_FUND_CODE' or col_upper == 'PARENT SUB FUND CODE':
                                parent_sub_fund_code_col = col
                            elif col_upper == 'ASSET_CLASS' or col_upper == 'ASSET CLASS' or ('ASSET' in col_upper and 'CLASS' in col_upper):
                                asset_class_col = col
                        except (TypeError, AttributeError):
                            # Skip columns that can't be converted to string or checked
                            continue
                    
                    if scheme_code_col and parent_sub_fund_code_col:
                        # Normalize RTA Master columns for matching
                        rta_df_normalized = rta_df.copy()
                        rta_df_normalized[scheme_code_col] = rta_df_normalized[scheme_code_col].astype(str).str.strip()
                        rta_df_normalized[parent_sub_fund_code_col] = rta_df_normalized[parent_sub_fund_code_col].astype(str).str.strip()
                        
                        # Create mapping: Scheme_code -> PARENT_SUB_FUND_CODE
                        subfund_mapping = rta_df_normalized.set_index(scheme_code_col)[parent_sub_fund_code_col].to_dict()
                        
                        # Create mapping: Scheme_code -> ASSET_CLASS (if column exists)
                        asset_class_mapping = {}
                        if asset_class_col:
                            rta_df_normalized[asset_class_col] = rta_df_normalized[asset_class_col].astype(str).str.strip()
                            asset_class_mapping = rta_df_normalized.set_index(scheme_code_col)[asset_class_col].to_dict()
                        
                        # Match "out Scheme Code" with Scheme_code and get PARENT_SUB_FUND_CODE
                        if 'out Scheme Code' in processed_df.columns:
                            processed_df['out Scheme Code'] = processed_df['out Scheme Code'].astype(str).str.strip()
                            processed_df['out subfund code'] = processed_df['out Scheme Code'].map(subfund_mapping)
                            processed_df['out subfund code'] = processed_df['out subfund code'].fillna('Not Found')
                            
                            # Get ASSET_CLASS for "out Scheme Code"
                            if asset_class_mapping:
                                processed_df['out ASSET_CLASS'] = processed_df['out Scheme Code'].map(asset_class_mapping)
                                processed_df['out ASSET_CLASS'] = processed_df['out ASSET_CLASS'].fillna('Not Found')
                            else:
                                processed_df['out ASSET_CLASS'] = 'Not Found'
                            
                            # Insert "out subfund code" right after "out Scheme Code"
                            cols = list(processed_df.columns)
                            out_scheme_idx = cols.index('out Scheme Code')
                            out_subfund = processed_df.pop('out subfund code')
                            out_asset_class = processed_df.pop('out ASSET_CLASS')
                            processed_df.insert(out_scheme_idx + 1, 'out subfund code', out_subfund)
                            processed_df.insert(out_scheme_idx + 2, 'out ASSET_CLASS', out_asset_class)
                        else:
                            processed_df['out subfund code'] = 'Not Found'
                            processed_df['out ASSET_CLASS'] = 'Not Found'
                        
                        # Match "IN Scheme Code" with Scheme_code and get PARENT_SUB_FUND_CODE
                        if 'IN Scheme Code' in processed_df.columns:
                            processed_df['IN Scheme Code'] = processed_df['IN Scheme Code'].astype(str).str.strip()
                            processed_df['IN subfund code'] = processed_df['IN Scheme Code'].map(subfund_mapping)
                            processed_df['IN subfund code'] = processed_df['IN subfund code'].fillna('Not Found')
                            
                            if asset_class_mapping:
                                processed_df['IN ASSET_CLASS'] = processed_df['IN Scheme Code'].map(asset_class_mapping)
                                processed_df['IN ASSET_CLASS'] = processed_df['IN ASSET_CLASS'].fillna('Not Found')
                            else:
                                processed_df['IN ASSET_CLASS'] = 'Not Found'
                            
                            cols = list(processed_df.columns)
                            in_scheme_idx = cols.index('IN Scheme Code')
                            in_subfund = processed_df.pop('IN subfund code')
                            in_asset_class = processed_df.pop('IN ASSET_CLASS')
                            processed_df.insert(in_scheme_idx + 1, 'IN subfund code', in_subfund)
                            processed_df.insert(in_scheme_idx + 2, 'IN ASSET_CLASS', in_asset_class)
                        else:
                            processed_df['IN subfund code'] = 'Not Found'
                            processed_df['IN ASSET_CLASS'] = 'Not Found'
                    else:
                        missing_cols = []
                        if not scheme_code_col:
                            missing_cols.append("Scheme_code")
                        if not parent_sub_fund_code_col:
                            missing_cols.append("PARENT_SUB_FUND_CODE")
                        
                        self.root.after(0, lambda: messagebox.showwarning(
                            "Warning",
                            f"Columns not found in RTA Master: {', '.join(missing_cols)}\n"
                            f"Subfund codes and ASSET_CLASS will not be added."
                        ))
                        processed_df['out subfund code'] = 'Not Found'
                        processed_df['out ASSET_CLASS'] = 'Not Found'
                        processed_df['IN subfund code'] = 'Not Found'
                        processed_df['IN ASSET_CLASS'] = 'Not Found'
                
                loading_window.update_status("Matching with Brokerage Structure...")
                
                # Match broker and IN subfund code with Brokerage Structure
                if not combined_brokerage_df.empty:
                    # DEBUG: Print Brokerage Structure columns
                    print("\n=== DEBUG: Brokerage Structure Columns ===")
                    print(list(combined_brokerage_df.columns))
                    print(f"Total rows: {len(combined_brokerage_df)}")
                    if len(combined_brokerage_df) > 0:
                        print("\nFirst few rows:")
                        print(combined_brokerage_df.head(3))
                    
                    # Find columns in Brokerage Structure
                    cons_code_col = None
                    scheme_code_b_col = None
                    trail_rate_cols = {}
                    investment_period_from_col = None
                    investment_period_to_col = None
                    
                    for col in combined_brokerage_df.columns:
                        try:
                            col_upper = str(col).upper().strip()
                            # More flexible matching for Cons Code
                            if 'CONS' in col_upper and 'CODE' in col_upper:
                                cons_code_col = col
                            elif col_upper == 'CONS_CODE' or col_upper == 'CONS':
                                cons_code_col = col
                            # More flexible matching for Scheme Code
                            elif 'SCHEME' in col_upper and 'CODE' in col_upper:
                                scheme_code_b_col = col
                            elif col_upper == 'SCHEME_CODE' or (col_upper == 'SCHEME' and scheme_code_b_col is None):
                                scheme_code_b_col = col
                            # Investment Period From
                            elif 'INVESTMENT' in col_upper and 'PERIOD' in col_upper and 'FROM' in col_upper:
                                investment_period_from_col = col
                            # Investment Period To
                            elif 'INVESTMENT' in col_upper and 'PERIOD' in col_upper and 'TO' in col_upper:
                                investment_period_to_col = col
                            # Trail Rate columns
                            elif 'TRAIL' in col_upper and '1' in col_upper and ('YEAR' in col_upper or 'YR' in col_upper):
                                trail_rate_cols[1] = col
                            elif 'TRAIL' in col_upper and '2' in col_upper and ('YEAR' in col_upper or 'YR' in col_upper):
                                trail_rate_cols[2] = col
                            elif 'TRAIL' in col_upper and '3' in col_upper and ('YEAR' in col_upper or 'YR' in col_upper):
                                trail_rate_cols[3] = col
                            elif 'TRAIL' in col_upper and '4' in col_upper and ('YEAR' in col_upper or 'YR' in col_upper):
                                trail_rate_cols[4] = col
                            elif 'TRAIL' in col_upper and '5' in col_upper and ('YEAR' in col_upper or 'YR' in col_upper):
                                trail_rate_cols[5] = col
                        except (TypeError, AttributeError):
                            # Skip columns that can't be converted to string or checked
                            continue
                    
                    # DEBUG: Print found columns
                    print(f"\n=== DEBUG: Found Columns ===")
                    print(f"Cons Code Column: {cons_code_col}")
                    print(f"Scheme Code Column: {scheme_code_b_col}")
                    print(f"Investment Period From Column: {investment_period_from_col}")
                    print(f"Investment Period To Column: {investment_period_to_col}")
                    print(f"Trail Rate Columns: {trail_rate_cols}")
                    
                    # Find broker column in Switch Register (flexible matching); use IN/OUT broker when available
                    broker_col = None
                    for col in processed_df.columns:
                        col_upper = str(col).upper().strip()
                        if 'BROK' in col_upper and ('DLR' in col_upper or 'DEALER' in col_upper):
                            broker_col = col
                            break
                        elif col_upper == 'BROKER' or col_upper == 'BROKER CODE' or col_upper == 'BROKER_CODE':
                            broker_col = col
                            break
                    if use_switch_columns:
                        broker_col = in_broker_col  # for switch-in trail matching
                        broker_col_out = out_broker_col  # for switch-out trail matching
                    else:
                        broker_col_out = broker_col
                    
                    print(f"Broker Column in Switch Register: {broker_col}")
                    print(f"Switch Register Columns: {list(processed_df.columns)}")
                    
                    # Find transaction date column in Switch Register (prefer OUT_TRADE_ from Switch Register)
                    tran_date_col = None
                    for col in processed_df.columns:
                        col_upper = str(col).upper().strip()
                        if col_upper in ('OUT_TRADE_', 'OUT_TRADE', 'OUT TRADE'):
                            tran_date_col = col
                            break
                        elif 'TRAN' in col_upper and 'DATE' in col_upper:
                            tran_date_col = col
                            break
                        elif col_upper == 'TRANSACTION DATE' or col_upper == 'TRANSACTION_DATE':
                            tran_date_col = col
                            break
                        elif col_upper == 'DATE':
                            tran_date_col = col
                            break
                    
                    print(f"Transaction Date Column: {tran_date_col}")
                    
                    if cons_code_col and scheme_code_b_col and broker_col:
                        # Normalize Brokerage Structure for matching
                        brokerage_normalized = combined_brokerage_df.copy()
                        brokerage_normalized[cons_code_col] = brokerage_normalized[cons_code_col].astype(str).str.strip().str.upper()
                        brokerage_normalized[scheme_code_b_col] = brokerage_normalized[scheme_code_b_col].astype(str).str.strip().str.upper()
                        
                        # Remove NaN values that might cause issues
                        brokerage_normalized = brokerage_normalized[
                            (brokerage_normalized[cons_code_col] != 'NAN') &
                            (brokerage_normalized[cons_code_col] != '') &
                            (brokerage_normalized[scheme_code_b_col] != 'NAN') &
                            (brokerage_normalized[scheme_code_b_col] != '')
                        ]
                        
                        # Convert Investment Period dates to datetime if columns exist
                        if investment_period_from_col and investment_period_from_col in brokerage_normalized.columns:
                            brokerage_normalized['_PERIOD_FROM_DT'] = pd.to_datetime(
                                brokerage_normalized[investment_period_from_col], 
                                errors='coerce',
                                dayfirst=True
                            )
                        else:
                            brokerage_normalized['_PERIOD_FROM_DT'] = pd.NaT
                        
                        if investment_period_to_col and investment_period_to_col in brokerage_normalized.columns:
                            brokerage_normalized['_PERIOD_TO_DT'] = pd.to_datetime(
                                brokerage_normalized[investment_period_to_col], 
                                errors='coerce',
                                dayfirst=True
                            )
                        else:
                            brokerage_normalized['_PERIOD_TO_DT'] = pd.NaT
                        
                        # Convert transaction date in processed_df if column exists
                        if tran_date_col and tran_date_col in processed_df.columns:
                            processed_df['_TRAN_DATE_DT'] = pd.to_datetime(
                                processed_df[tran_date_col], 
                                errors='coerce',
                                dayfirst=True
                            )
                        else:
                            processed_df['_TRAN_DATE_DT'] = pd.NaT
                        
                        # Build lookup: (broker, scheme) -> subset of brokerage for fast per-row access
                        brokerage_lookup = {}
                        for (b, s), grp in brokerage_normalized.groupby([cons_code_col, scheme_code_b_col], dropna=False):
                            try:
                                k = (str(b).strip().upper() if pd.notna(b) else '', str(s).strip().upper() if pd.notna(s) else '')
                                if k[0] != '' and k[1] != '':
                                    brokerage_lookup[k] = grp
                            except (TypeError, AttributeError):
                                pass
                        empty_df = pd.DataFrame()
                        
                        # DEBUG: Print sample values
                        print(f"\n=== DEBUG: Sample Brokerage Values ===")
                        if len(brokerage_normalized) > 0:
                            print(f"Sample Cons Code values: {brokerage_normalized[cons_code_col].head(5).tolist()}")
                            print(f"Sample Scheme Code values: {brokerage_normalized[scheme_code_b_col].head(5).tolist()}")
                            if investment_period_from_col:
                                print(f"Sample Period From values: {brokerage_normalized[investment_period_from_col].head(3).tolist()}")
                            if investment_period_to_col:
                                print(f"Sample Period To values: {brokerage_normalized[investment_period_to_col].head(3).tolist()}")
                        
                        # Create a function to get trail rates
                        match_count = 0
                        no_match_count = 0
                        
                        def get_trail_rates_and_periods(row):
                            nonlocal match_count, no_match_count
                            try:
                                # Get broker code
                                broker = row.get(broker_col) if broker_col in row.index else None
                                if pd.isna(broker) or broker == '':
                                    no_match_count += 1
                                    return ['Not Found'] * 3  # 1 trail rate + 2 period columns
                                
                                broker_str = str(broker).strip().upper()
                                
                                # Get IN subfund code
                                in_subfund = row.get('IN subfund code') if 'IN subfund code' in row.index else None
                                if pd.isna(in_subfund) or in_subfund == 'Not Found' or in_subfund == '':
                                    no_match_count += 1
                                    return ['Not Found'] * 3
                                
                                in_subfund_str = str(in_subfund).strip().upper()
                                
                                # Get transaction date
                                tran_date = row.get('_TRAN_DATE_DT') if '_TRAN_DATE_DT' in row.index else None
                                
                                # Calculate previous month from transaction date
                                previous_month_date = None
                                if pd.notna(tran_date):
                                    try:
                                        # Get the first day of the transaction month
                                        first_day_current = tran_date.replace(day=1)
                                        # Subtract one day to get last day of previous month
                                        last_day_previous = first_day_current - pd.Timedelta(days=1)
                                        # Get first day of previous month
                                        previous_month_date = last_day_previous.replace(day=1)
                                    except Exception:
                                        previous_month_date = None
                                
                                # DEBUG: Print first few attempts only (avoid I/O slowdown)
                                if match_count + no_match_count < 2:
                                    print(f"\n=== DEBUG: Matching Attempt {match_count + no_match_count + 1} ===")
                                    print(f"Broker: '{broker_str}'")
                                    print(f"IN Subfund: '{in_subfund_str}'")
                                    print(f"Transaction Date: {tran_date}")
                                    print(f"Previous Month Date: {previous_month_date}")
                                
                                # Look up brokerage rows for this (broker, scheme) — O(1) instead of full scan
                                matches = brokerage_lookup.get((broker_str, in_subfund_str), empty_df)
                                
                                if matches.empty:
                                    no_match_count += 1
                                    if match_count + no_match_count <= 2:
                                        print(f"  No match for (broker, IN subfund)")
                                    return ['Not Found'] * 4  # 1 current + 1 previous trail rate + 2 period columns
                                
                                # Vectorized date match: find row(s) where tran_date in [period_from, period_to]
                                current_rate_match = None
                                previous_rate_match = None
                                pf = matches['_PERIOD_FROM_DT']
                                pt = matches['_PERIOD_TO_DT']
                                has_period = pf.notna() & pt.notna()
                                
                                if pd.notna(tran_date) and investment_period_from_col and investment_period_to_col and has_period.any():
                                    # Current: first row where period contains tran_date
                                    in_range_cur = (pf <= tran_date) & (tran_date <= pt)
                                    if in_range_cur.any():
                                        current_rate_match = matches.loc[in_range_cur].iloc[0]
                                        if match_count + no_match_count < 2:
                                            print(f"  ✓ Current month match found!")
                                    
                                    # Previous month (same year): first row where period contains previous_month_date
                                    if previous_month_date and pd.notna(previous_month_date) and previous_month_date.year == tran_date.year:
                                        in_range_prev = (pf <= previous_month_date) & (previous_month_date <= pt)
                                        if in_range_prev.any():
                                            previous_rate_match = matches.loc[in_range_prev].iloc[0]
                                            if match_count + no_match_count < 2:
                                                print(f"  ✓ Previous month match found! (same year)")
                                else:
                                    # No date filtering - use first match for current only
                                    if not matches.empty:
                                        current_rate_match = matches.iloc[0]
                                        if match_count + no_match_count < 2:
                                            print(f"  ⚠ Using first match (no date filter)")
                                
                                match_count += 1
                                if match_count <= 2:
                                    print(f"  ✓ Match found!")
                                
                                results = []
                                
                                # Get Trail Rate 1 year for CURRENT month
                                if current_rate_match is not None and 1 in trail_rate_cols:
                                    trail_col = trail_rate_cols[1]
                                    if trail_col in current_rate_match.index:
                                        rate_value = current_rate_match[trail_col]
                                        if pd.notna(rate_value) and str(rate_value).strip() != '':
                                            results.append(rate_value)
                                        else:
                                            results.append('Not Found')
                                    else:
                                        results.append('Not Found')
                                else:
                                    results.append('Not Found')
                                
                                # Get Trail Rate 1 year for PREVIOUS month
                                if previous_rate_match is not None and 1 in trail_rate_cols:
                                    trail_col = trail_rate_cols[1]
                                    if trail_col in previous_rate_match.index:
                                        rate_value = previous_rate_match[trail_col]
                                        if pd.notna(rate_value) and str(rate_value).strip() != '':
                                            results.append(rate_value)
                                        else:
                                            results.append('Not Found')
                                    else:
                                        results.append('Not Found')
                                else:
                                    results.append('Not Found')
                                
                                # Get Investment Period From (from current match)
                                if current_rate_match is not None and investment_period_from_col and investment_period_from_col in current_rate_match.index:
                                    period_from = current_rate_match[investment_period_from_col]
                                    if pd.notna(period_from) and str(period_from).strip() != '':
                                        results.append(period_from)
                                    else:
                                        results.append('Not Found')
                                else:
                                    results.append('Not Found')
                                
                                # Get Investment Period To (from current match)
                                if current_rate_match is not None and investment_period_to_col and investment_period_to_col in current_rate_match.index:
                                    period_to = current_rate_match[investment_period_to_col]
                                    if pd.notna(period_to) and str(period_to).strip() != '':
                                        results.append(period_to)
                                    else:
                                        results.append('Not Found')
                                else:
                                    results.append('Not Found')
                                
                                return results
                            except Exception as e:
                                no_match_count += 1
                                if match_count + no_match_count <= 2:
                                    print(f"  Exception: {str(e)}")
                                return ['Not Found'] * 4
                        
                        # Apply trail rate and period extraction
                        trail_results = processed_df.apply(get_trail_rates_and_periods, axis=1, result_type='expand')
                        trail_results.columns = ['switch in Trail Rate 1 year', 'PREVIOUS switch in Trail Rate 1 year', 'Investment Period From', 'Investment Period To']
                        
                        # Add trail rate columns to processed_df
                        processed_df['switch in Trail Rate 1 year'] = trail_results['switch in Trail Rate 1 year']
                        processed_df['PREVIOUS switch in Trail Rate 1 year'] = trail_results['PREVIOUS switch in Trail Rate 1 year']
                        
                        # Add Investment Period columns
                        processed_df['Investment Period From'] = trail_results['Investment Period From']
                        processed_df['Investment Period To'] = trail_results['Investment Period To']
                        
                        # DEBUG: Print summary
                        print(f"\n=== DEBUG: Matching Summary (IN) ===")
                        print(f"Total matches: {match_count}")
                        print(f"Total no matches: {no_match_count}")
                        print(f"Total rows processed: {len(processed_df)}")
                        
                        # Now match with OUT subfund code for "switch out" trail rates
                        loading_window.update_status("Matching with Brokerage Structure (OUT)...")
                        
                        match_count_out = 0
                        no_match_count_out = 0
                        
                        def get_trail_rates_out(row):
                            nonlocal match_count_out, no_match_count_out
                            try:
                                # Get broker code (OUT broker when use_switch_columns)
                                broker = row.get(broker_col_out) if broker_col_out in row.index else None
                                if pd.isna(broker) or broker == '':
                                    no_match_count_out += 1
                                    return ['Not Found']  # 1 trail rate only
                                
                                broker_str = str(broker).strip().upper()
                                
                                # Get OUT subfund code
                                out_subfund = row.get('out subfund code') if 'out subfund code' in row.index else None
                                if pd.isna(out_subfund) or out_subfund == 'Not Found' or out_subfund == '':
                                    no_match_count_out += 1
                                    return ['Not Found']
                                
                                out_subfund_str = str(out_subfund).strip().upper()
                                
                                # Get transaction date
                                tran_date = row.get('_TRAN_DATE_DT') if '_TRAN_DATE_DT' in row.index else None
                                
                                # Calculate previous month from transaction date
                                previous_month_date = None
                                if pd.notna(tran_date):
                                    try:
                                        # Get the first day of the transaction month
                                        first_day_current = tran_date.replace(day=1)
                                        # Subtract one day to get last day of previous month
                                        last_day_previous = first_day_current - pd.Timedelta(days=1)
                                        # Get first day of previous month
                                        previous_month_date = last_day_previous.replace(day=1)
                                    except Exception:
                                        previous_month_date = None
                                
                                # Look up brokerage rows for this (broker, scheme) — O(1) instead of full scan
                                matches = brokerage_lookup.get((broker_str, out_subfund_str), empty_df)
                                
                                if matches.empty:
                                    no_match_count_out += 1
                                    return ['Not Found', 'Not Found']  # Current + Previous
                                
                                current_rate_match = None
                                previous_rate_match = None
                                # Vectorized date match (same as IN)
                                pf = matches['_PERIOD_FROM_DT']
                                pt = matches['_PERIOD_TO_DT']
                                has_period = pf.notna() & pt.notna()
                                
                                if pd.notna(tran_date) and investment_period_from_col and investment_period_to_col and has_period.any():
                                    in_range_cur = (pf <= tran_date) & (tran_date <= pt)
                                    if in_range_cur.any():
                                        current_rate_match = matches.loc[in_range_cur].iloc[0]
                                    if previous_month_date and pd.notna(previous_month_date) and previous_month_date.year == tran_date.year:
                                        in_range_prev = (pf <= previous_month_date) & (previous_month_date <= pt)
                                        if in_range_prev.any():
                                            previous_rate_match = matches.loc[in_range_prev].iloc[0]
                                else:
                                    if not matches.empty:
                                        current_rate_match = matches.iloc[0]
                                
                                match_count_out += 1
                                
                                results = []
                                
                                # Get Trail Rate 1 year for CURRENT month
                                if current_rate_match is not None and 1 in trail_rate_cols:
                                    trail_col = trail_rate_cols[1]
                                    if trail_col in current_rate_match.index:
                                        rate_value = current_rate_match[trail_col]
                                        if pd.notna(rate_value) and str(rate_value).strip() != '':
                                            results.append(rate_value)
                                        else:
                                            results.append('Not Found')
                                    else:
                                        results.append('Not Found')
                                else:
                                    results.append('Not Found')
                                
                                # Get Trail Rate 1 year for PREVIOUS month
                                if previous_rate_match is not None and 1 in trail_rate_cols:
                                    trail_col = trail_rate_cols[1]
                                    if trail_col in previous_rate_match.index:
                                        rate_value = previous_rate_match[trail_col]
                                        if pd.notna(rate_value) and str(rate_value).strip() != '':
                                            results.append(rate_value)
                                        else:
                                            results.append('Not Found')
                                    else:
                                        results.append('Not Found')
                                else:
                                    results.append('Not Found')
                                
                                return results
                            except Exception as e:
                                no_match_count_out += 1
                                return ['Not Found', 'Not Found']
                        
                        # Apply trail rate extraction for OUT subfund code
                        trail_results_out = processed_df.apply(get_trail_rates_out, axis=1, result_type='expand')
                        trail_results_out.columns = ['switch out Trail Rate 1 year', 'PREVIOUS switch out Trail Rate 1 year']
                        
                        # Add "switch out" trail rate columns to processed_df
                        processed_df['switch out Trail Rate 1 year'] = trail_results_out['switch out Trail Rate 1 year']
                        processed_df['PREVIOUS switch out Trail Rate 1 year'] = trail_results_out['PREVIOUS switch out Trail Rate 1 year']
                        
                        # DEBUG: Print summary for OUT
                        print(f"\n=== DEBUG: Matching Summary (OUT) ===")
                        print(f"Total matches: {match_count_out}")
                        print(f"Total no matches: {no_match_count_out}")
                        
                        # Add check columns: if switch in > switch out, show "Check"
                        loading_window.update_status("Calculating checks...")
                        
                        def compare_trail_rates(value_in, value_out):
                            """Compare two trail rate values and return 'Check' if in > out"""
                            try:
                                # Handle "Not Found" or empty values
                                if pd.isna(value_in) or str(value_in).strip() == '' or str(value_in).strip().upper() == 'NOT FOUND':
                                    return ''
                                if pd.isna(value_out) or str(value_out).strip() == '' or str(value_out).strip().upper() == 'NOT FOUND':
                                    return ''
                                
                                # Convert to numeric, handling percentage signs and other formats
                                val_in_str = str(value_in).strip().replace('%', '').replace(',', '')
                                val_out_str = str(value_out).strip().replace('%', '').replace(',', '')
                                
                                try:
                                    val_in_num = float(val_in_str)
                                    val_out_num = float(val_out_str)
                                    
                                    # If switch in > switch out, return "Check"
                                    if val_in_num > val_out_num:
                                        return 'Check'
                                    else:
                                        return ''
                                except (ValueError, TypeError):
                                    # If conversion fails, return empty
                                    return ''
                            except Exception:
                                return ''
                        
                        # Add check column for 1 year only
                        col_in = 'switch in Trail Rate 1 year'
                        col_out = 'switch out Trail Rate 1 year'
                        check_col = 'Check 1 year'
                        
                        if col_in in processed_df.columns and col_out in processed_df.columns:
                            processed_df[check_col] = processed_df.apply(
                                lambda row: compare_trail_rates(row[col_in], row[col_out]),
                                axis=1
                            )
                        else:
                            processed_df[check_col] = ''
                        
                        # Add check: PREVIOUS switch in Trail Rate 1 year < switch in Trail Rate 1 year
                        loading_window.update_status("Calculating previous vs current check...")
                        
                        def compare_previous_vs_current(previous_value, current_value):
                            """Compare previous and current trail rate values and return 'Check' if previous < current"""
                            try:
                                # Handle "Not Found" or empty values
                                if pd.isna(previous_value) or str(previous_value).strip() == '' or str(previous_value).strip().upper() == 'NOT FOUND':
                                    return ''
                                if pd.isna(current_value) or str(current_value).strip() == '' or str(current_value).strip().upper() == 'NOT FOUND':
                                    return ''
                                
                                # Convert to numeric, handling percentage signs and other formats
                                prev_str = str(previous_value).strip().replace('%', '').replace(',', '')
                                curr_str = str(current_value).strip().replace('%', '').replace(',', '')
                                
                                try:
                                    prev_num = float(prev_str)
                                    curr_num = float(curr_str)
                                    
                                    # If previous < current, return "Check"
                                    if prev_num < curr_num:
                                        return 'Check'
                                    else:
                                        return ''
                                except (ValueError, TypeError):
                                    # If conversion fails, return empty
                                    return ''
                            except Exception:
                                return ''
                        
                        # Add check column for previous vs current
                        prev_col = 'PREVIOUS switch in Trail Rate 1 year'
                        curr_col = 'switch in Trail Rate 1 year'
                        prev_check_col = 'PREVIOUS switch in Trail Rate 1 year Check'
                        
                        if prev_col in processed_df.columns and curr_col in processed_df.columns:
                            processed_df[prev_check_col] = processed_df.apply(
                                lambda row: compare_previous_vs_current(row[prev_col], row[curr_col]),
                                axis=1
                            )
                        else:
                            processed_df[prev_check_col] = ''
                    else:
                        # Missing required columns
                        missing_cols = []
                        if not cons_code_col:
                            missing_cols.append("Cons Code")
                        if not scheme_code_b_col:
                            missing_cols.append("Scheme Code")
                        if not broker_col:
                            missing_cols.append("Broker Column (BROK_DLR_N or similar)")
                        
                        self.root.after(0, lambda m=missing_cols: messagebox.showwarning(
                            "Warning",
                            f"Columns not found:\n"
                            f"Brokerage Structure: {', '.join([c for c in m if c != 'Broker Column (BROK_DLR_N or similar)'])}\n"
                            f"Switch Register: {', '.join([c for c in m if c == 'Broker Column (BROK_DLR_N or similar)'])}\n"
                            f"Trail rates will not be added."
                        ))
                        # Add empty trail rate columns
                        processed_df['switch in Trail Rate 1 year'] = 'Not Found'
                        processed_df['PREVIOUS switch in Trail Rate 1 year'] = 'Not Found'
                        processed_df['switch out Trail Rate 1 year'] = 'Not Found'
                        processed_df['PREVIOUS switch out Trail Rate 1 year'] = 'Not Found'
                        processed_df['Check 1 year'] = ''
                        processed_df['PREVIOUS switch in Trail Rate 1 year Check'] = ''
                        processed_df['Investment Period From'] = 'Not Found'
                        processed_df['Investment Period To'] = 'Not Found'
                else:
                    # No brokerage structure data
                    processed_df['switch in Trail Rate 1 year'] = 'Not Found'
                    processed_df['PREVIOUS switch in Trail Rate 1 year'] = 'Not Found'
                    processed_df['switch out Trail Rate 1 year'] = 'Not Found'
                    processed_df['PREVIOUS switch out Trail Rate 1 year'] = 'Not Found'
                    processed_df['Check 1 year'] = ''
                    processed_df['PREVIOUS switch in Trail Rate 1 year Check'] = ''
                    processed_df['Investment Period From'] = 'Not Found'
                    processed_df['Investment Period To'] = 'Not Found'
                
                # Reorder columns: Group by year (switch in, switch out, check for each year)
                loading_window.update_status("Reordering columns...")
                
                # Get all current columns
                all_columns = list(processed_df.columns)
                
                # Find trail rate and check columns
                switch_in_col = None
                previous_switch_in_col = None
                switch_out_col = None
                previous_switch_out_col = None
                check_col = None
                prev_check_col = None
                
                for col in all_columns:
                    try:
                        # Ensure col is a string before using 'in' operator
                        col_str = str(col) if not isinstance(col, str) else col
                        
                        if col_str == 'switch in Trail Rate 1 year':
                            switch_in_col = col
                        elif col_str == 'PREVIOUS switch in Trail Rate 1 year':
                            previous_switch_in_col = col
                        elif col_str == 'switch out Trail Rate 1 year':
                            switch_out_col = col
                        elif col_str == 'PREVIOUS switch out Trail Rate 1 year':
                            previous_switch_out_col = col
                        elif col_str == 'Check 1 year':
                            check_col = col
                        elif col_str == 'PREVIOUS switch in Trail Rate 1 year Check':
                            prev_check_col = col
                    except (TypeError, AttributeError):
                        # Skip columns that can't be converted to string
                        pass
                
                # Build new column order: switch in, previous switch in, previous check, switch out, previous switch out, check
                new_column_order = []
                
                # Add other columns first (before trail rate columns)
                for col in all_columns:
                    if col not in [switch_in_col, previous_switch_in_col, switch_out_col, previous_switch_out_col, check_col, prev_check_col]:
                        if col not in new_column_order:
                            new_column_order.append(col)
                
                # Add trail rate columns in order: switch in, previous switch in, previous check, switch out, previous switch out, check
                if switch_in_col:
                    new_column_order.append(switch_in_col)
                if previous_switch_in_col:
                    new_column_order.append(previous_switch_in_col)
                if prev_check_col:
                    new_column_order.append(prev_check_col)
                if switch_out_col:
                    new_column_order.append(switch_out_col)
                if previous_switch_out_col:
                    new_column_order.append(previous_switch_out_col)
                if check_col:
                    new_column_order.append(check_col)
                
                # Reorder the dataframe
                processed_df = processed_df[new_column_order]
                
                # Add check for Regular vs Direct scheme matching
                loading_window.update_status("Checking Regular vs Direct scheme matches...")
                
                # Find switch in/out scheme columns (IN_SCHEME_, OUT_SCHEM0/OUT_SCHEME or legacy)
                def _norm_col(s):
                    try:
                        t = str(s).upper().lstrip('\ufeff').strip()
                        return ' '.join(t.split()).replace(' ', '_')
                    except (TypeError, AttributeError):
                        return ''
                in_scheme_col = None
                out_scheme_col = None
                for col in processed_df.columns:
                    col_str = str(col).strip()
                    c = _norm_col(col)
                    # Exact match first (user's columns: IN_SCHEME_, OUT_SCHEM0)
                    if col_str.upper() in ('IN_SCHEME_', 'IN_SCHEME'):
                        in_scheme_col = col
                    if col_str.upper() in ('OUT_SCHEM0', 'OUT_SCHEME'):
                        out_scheme_col = col
                if not in_scheme_col:
                    for col in processed_df.columns:
                        c = _norm_col(col)
                        if 'IN_SCHEME' in c or c in ('IN_SCHEME_', 'INSCHEME', 'SCHEME_IN'):
                            in_scheme_col = col
                            break
                if not out_scheme_col:
                    for col in processed_df.columns:
                        c = _norm_col(col)
                        if 'OUT_SCHEM' in c or 'OUT_SCHEME' in c or c in ('OUTSCHEME', 'SCHEME_OUT'):
                            out_scheme_col = col
                            break
                if not in_scheme_col and 'switch in scheme' in processed_df.columns:
                    in_scheme_col = 'switch in scheme'
                if not out_scheme_col and 'switch out scheme' in processed_df.columns:
                    out_scheme_col = 'switch out scheme'
                
                def check_regular_direct_match(row):
                    """Check if switch in and switch out schemes are the same but one is Regular and other is Direct"""
                    try:
                        switch_in = row.get(in_scheme_col) if in_scheme_col and in_scheme_col in row.index else None
                        switch_out = row.get(out_scheme_col) if out_scheme_col and out_scheme_col in row.index else None
                        
                        if pd.isna(switch_in) or pd.isna(switch_out):
                            return ''
                        
                        # Ensure we have strings
                        try:
                            switch_in_str = str(switch_in).strip() if switch_in is not None else ''
                            switch_out_str = str(switch_out).strip() if switch_out is not None else ''
                        except Exception:
                            return ''
                        
                        if switch_in_str == '' or switch_out_str == '':
                            return ''
                        
                        # Remove code prefix (e.g., "129B/ABSL" or any code before "/")
                        def remove_code_prefix(scheme_str):
                            try:
                                scheme_str = str(scheme_str) if not isinstance(scheme_str, str) else scheme_str
                                if '/' in scheme_str:
                                    # Split by "/" and take everything after the first "/"
                                    parts = scheme_str.split('/', 1)
                                    if len(parts) > 1:
                                        return parts[1].strip()
                                return scheme_str
                            except Exception:
                                return str(scheme_str) if scheme_str is not None else ''
                        
                        scheme_in = remove_code_prefix(switch_in_str)
                        scheme_out = remove_code_prefix(switch_out_str)
                        
                        if not scheme_in or not scheme_out:
                            return ''
                        
                        # Normalize: remove Regular, Direct, Reg, Growth, Plan so "Regular" vs "Direct" don't hurt similarity
                        def normalize_for_match(scheme):
                            try:
                                s = str(scheme).upper()
                                for word in ['REGULAR GROWTH', 'REG GROWTH', 'REGULAR', 'REG PLAN', 'DIRECT GROWTH', 'DIRECT', 'REG', 'GROWTH', 'PLAN']:
                                    s = s.replace(word, '')
                                s = s.replace('-', ' ').replace('_', ' ').strip()
                                while '  ' in s:
                                    s = s.replace('  ', ' ')
                                return s.strip()
                            except Exception:
                                return str(scheme) if scheme else ''
                        
                        norm_in = normalize_for_match(scheme_in)
                        norm_out = normalize_for_match(scheme_out)
                        if not norm_in and not norm_out:
                            return ''
                        if not norm_in or not norm_out:
                            return ''
                        # Exact match or fuzzy 80% on normalized base names
                        if norm_in == norm_out:
                            pass  # same scheme
                        else:
                            similarity = difflib.SequenceMatcher(None, norm_in, norm_out).ratio()
                            if similarity < 0.75:
                                return ''
                        
                        # Check if one has Regular and other has Direct
                        try:
                            switch_in_upper = str(switch_in_str).upper() if switch_in_str is not None else ''
                            switch_out_upper = str(switch_out_str).upper() if switch_out_str is not None else ''
                            if not isinstance(switch_in_upper, str) or not isinstance(switch_out_upper, str):
                                return ''
                            # Regular: REGULAR, REG GROWTH, REG PLAN, REG (with word boundary)
                            has_regular_in = 'REGULAR' in switch_in_upper or 'REG GROWTH' in switch_in_upper or 'REG PLAN' in switch_in_upper or '-REG-' in switch_in_upper or ' REG ' in switch_in_upper
                            has_regular_out = 'REGULAR' in switch_out_upper or 'REG GROWTH' in switch_out_upper or 'REG PLAN' in switch_out_upper or '-REG-' in switch_out_upper or ' REG ' in switch_out_upper
                            # Direct
                            has_direct_in = 'DIRECT' in switch_in_upper
                            has_direct_out = 'DIRECT' in switch_out_upper
                        except (TypeError, AttributeError):
                            return ''
                        # One Regular + one Direct = Check (e.g. -PLAN-Regular-Growth vs -PLAN-Direct-Growth)
                        if (has_regular_in and has_direct_out) or (has_direct_in and has_regular_out):
                            return 'Check'
                        
                        return ''
                    except Exception as e:
                        return ''
                
                # Add the check column
                if in_scheme_col and out_scheme_col and in_scheme_col in processed_df.columns and out_scheme_col in processed_df.columns:
                    processed_df['Regular vs Direct Check'] = processed_df.apply(check_regular_direct_match, axis=1)
                else:
                    processed_df['Regular vs Direct Check'] = ''
                    if not in_scheme_col or not out_scheme_col:
                        cols_found = [c for c in processed_df.columns if 'SCHEME' in str(c).upper() or 'SCHEM' in str(c).upper()]
                        self.root.after(0, lambda: messagebox.showwarning(
                            "Regular vs Direct Check",
                            f"Could not find scheme columns for Regular vs Direct check.\n"
                            f"Looking for: IN_SCHEME_ (or IN_SCHEME), OUT_SCHEM0 (or OUT_SCHEME)\n"
                            f"Columns with SCHEME in name: {cols_found[:10] if cols_found else 'None'}"
                        ))
                
                loading_window.update_status("Processing complete...")
                time.sleep(0.3)
                
                # Remove rows where broker is "DIRECT"
                loading_window.update_status("Filtering out DIRECT broker entries...")
                
                # Find broker column for DIRECT filter (use OUT_BROKER when we have IN/OUT columns)
                broker_col_filter = None
                if use_switch_columns and out_broker_col and out_broker_col in processed_df.columns:
                    broker_col_filter = out_broker_col
                else:
                    for col in processed_df.columns:
                        col_upper = str(col).upper().strip()
                        if 'BROK' in col_upper and ('DLR' in col_upper or 'DEALER' in col_upper):
                            broker_col_filter = col
                            break
                        elif col_upper == 'BROKER' or col_upper == 'BROKER CODE' or col_upper == 'BROKER_CODE':
                            broker_col_filter = col
                            break
                
                if broker_col_filter and broker_col_filter in processed_df.columns:
                    initial_count = len(processed_df)
                    # Filter out rows where broker is "DIRECT" (case-insensitive)
                    processed_df = processed_df[
                        ~processed_df[broker_col_filter].astype(str).str.strip().str.upper().eq('DIRECT')
                    ]
                    removed_count = initial_count - len(processed_df)
                    if removed_count > 0:
                        print(f"\n=== Removed {removed_count} rows with DIRECT broker ===")
                        self.root.after(0, lambda: messagebox.showinfo(
                            "Filter Applied",
                            f"Removed {removed_count} row(s) where broker is 'DIRECT'.\n"
                            f"Remaining rows: {len(processed_df)}"
                        ))
                else:
                    print("\n=== Warning: Broker column not found, skipping DIRECT filter ===")
                
                # Output only these columns, in this exact order
                OUTPUT_COLUMNS = [
                    'SL_NO', 'FOLIO_NO', 'INVESTOR_F', 'USER_TRXNN', 'OUT_TRXN_N',
                    'SO_ASSET_C', 'OUT_SUBFUN', 'OUT_SCHEME', 'OUT_SCHEM0', 'OUT_TRADE_',
                    'OUT_BROKER', 'OUT_BROKE1', 'SO_UNITS', 'SO_AMOUNT',
                    'IN_SUBFUND', 'IN_SCHEME', 'IN_SCHEME_', 'SI_ASSET_C', 'IN_TRADE_D',
                    'IN_BROKER', 'SI_UNITS', 'SI_AMOUNT',
                    'IN ASSET_CLASS', 'out ASSET_CLASS',
                    'Investment Period From', 'Investment Period To',
                    'switch in Trail Rate 1 year', 'PREVIOUS switch in Trail Rate 1 year',
                    'PREVIOUS switch in Trail Rate 1 year Check',
                    'switch out Trail Rate 1 year', 'PREVIOUS switch out Trail Rate 1 year',
                    'Check 1 year', 'Regular vs Direct Check'
                ]
                final_cols = []
                for col in OUTPUT_COLUMNS:
                    if col in processed_df.columns:
                        final_cols.append(col)
                    else:
                        processed_df[col] = ''
                        final_cols.append(col)
                processed_df = processed_df[final_cols]
                
                loading_window.update_status("Preparing to save...")
                save_path = filedialog.asksaveasfilename(
                    title="Save Processed File",
                    defaultextension=".xlsx",
                    filetypes=[
                        ("Excel files", "*.xlsx"),
                        ("CSV files (faster for large files)", "*.csv"),
                        ("All files", "*.*")
                    ]
                )
                
                if save_path:
                    try:
                        loading_window.update_status("Saving file...")
                        # CSV: very fast even for millions of rows (recommended for large files)
                        if save_path.lower().endswith('.csv'):
                            processed_df.to_csv(save_path, index=False, encoding='utf-8-sig')
                        else:
                            # Excel: can be slow for 20k+ rows — use CSV for large files
                            try:
                                processed_df.to_excel(save_path, sheet_name='Processed Data', index=False, engine='xlsxwriter')
                            except ImportError:
                                processed_df.to_excel(save_path, sheet_name='Processed Data', index=False, engine='openpyxl')
                        
                        loading_window.update_status("Complete!")
                        time.sleep(0.5)
                        
                        self.root.after(0, lambda: messagebox.showinfo(
                            "Success",
                            f"File processed and saved successfully at:\n{save_path}"
                        ))
                        self.root.after(0, lambda: self.status_label.configure(
                            text="Processing completed successfully!",
                            text_color="#27ae60"
                        ))
                    except Exception as e:
                        self.root.after(0, lambda: messagebox.showerror(
                            "Error",
                            f"Could not save file:\n{e}"
                        ))
                        self.root.after(0, lambda: self.status_label.configure(
                            text="Status: Could not save file!",
                            text_color="#e74c3c"
                        ))
                else:
                    self.root.after(0, lambda: messagebox.showinfo(
                        "Cancelled",
                        "File save was cancelled."
                    ))
                    self.root.after(0, lambda: self.status_label.configure(
                        text="Status: File save was cancelled.",
                        text_color="#f39c12"
                    ))
                
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror(
                    "Error",
                    f"Error processing files: {str(e)}"
                ))
                self.root.after(0, lambda: self.status_label.configure(
                    text=f"Status: Error processing files: {str(e)}",
                    text_color="#e74c3c"
                ))
            finally:
                self.root.after(0, loading_window.stop)
        
        # Run processing in a separate thread
        thread = threading.Thread(target=process_file)
        thread.daemon = True
        thread.start()


def main():
    root = ctk.CTk()
    app = SwitchRegisterGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()

