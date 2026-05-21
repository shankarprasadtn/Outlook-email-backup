import os
import sys
import threading
import time
import logging
import tkinter as tk
from tkinter import filedialog, messagebox

# Fallback mechanism if CustomTkinter is not installed yet
try:
    import customtkinter as ctk
    ctk_available = True
except ImportError:
    ctk_available = False

import config
import sync_engine

class OutlookSyncApp:
    def __init__(self, root):
        self.root = root
        self.config_data = config.load_config()
        self.is_syncing = False
        self.scheduler_running = False
        self.scheduler_thread = None
        self.scheduler_stop_event = None
        
        # Windows title and dimensions
        self.root.title("M365 Outlook & SharePoint Sync Utility")
        self.root.geometry("820x600")
        self.root.minsize(800, 550)
        
        # Configure Grid layout
        self.root.grid_columnconfigure(1, weight=1)
        self.root.grid_rowconfigure(0, weight=1)
        
        # Styling parameters
        self.accent_color = "#0078d4" # Microsoft Blue
        self.bg_color_sidebar = "#202020" if ctk_available else "#2f3136"
        
        # Initialize UI components
        if ctk_available:
            ctk.set_appearance_mode("Dark")
            ctk.set_default_color_theme("blue")
            self.setup_ctk_ui()
        else:
            self.setup_standard_ui()
            
        # Hook sync_engine logs to our GUI log console
        sync_engine.setup_gui_logging(self.append_log)
        
        self.append_log("Application initialized successfully.", logging.INFO)
        self.append_log("Ready to synchronize.", logging.INFO)
        
        # Auto-start scheduler if active in config (could be added as default running)
        self.start_scheduler()
        
    def setup_ctk_ui(self):
        # Sidebar Frame
        self.sidebar_frame = ctk.CTkFrame(self.root, width=170, corner_radius=0, fg_color="#181818")
        self.sidebar_frame.grid(row=0, column=0, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(5, weight=1)
        
        # Sidebar Title
        self.title_label = ctk.CTkLabel(
            self.sidebar_frame, 
            text="Outlook Sync", 
            font=ctk.CTkFont(size=18, weight="bold"),
            text_color="#ffffff"
        )
        self.title_label.grid(row=0, column=0, padx=20, pady=(20, 30))
        
        # Navigation Buttons
        self.dash_btn = ctk.CTkButton(
            self.sidebar_frame, 
            text="Dashboard", 
            fg_color="transparent", 
            text_color="#d0d0d0",
            hover_color="#282828",
            anchor="w",
            command=self.show_dashboard
        )
        self.dash_btn.grid(row=1, column=0, padx=10, pady=5, sticky="ew")
        
        self.settings_btn = ctk.CTkButton(
            self.sidebar_frame, 
            text="Configuration", 
            fg_color="transparent", 
            text_color="#d0d0d0",
            hover_color="#282828",
            anchor="w",
            command=self.show_settings
        )
        self.settings_btn.grid(row=2, column=0, padx=10, pady=5, sticky="ew")
        
        self.logs_btn = ctk.CTkButton(
            self.sidebar_frame, 
            text="Console Logs", 
            fg_color="transparent", 
            text_color="#d0d0d0",
            hover_color="#282828",
            anchor="w",
            command=self.show_logs
        )
        self.logs_btn.grid(row=3, column=0, padx=10, pady=5, sticky="ew")
        
        # Status Badge inside Sidebar
        self.sched_badge = ctk.CTkLabel(
            self.sidebar_frame, 
            text="Scheduler: ACTIVE", 
            font=ctk.CTkFont(size=10, weight="bold"),
            text_color="#107c41"
        )
        self.sched_badge.grid(row=6, column=0, padx=10, pady=20)
        
        # Content Frames Container
        self.content_frame = ctk.CTkFrame(self.root, fg_color="#202020", corner_radius=0)
        self.content_frame.grid(row=0, column=1, sticky="nsew")
        self.content_frame.grid_columnconfigure(0, weight=1)
        self.content_frame.grid_rowconfigure(0, weight=1)
        
        # Build Pages
        self.create_dashboard_page()
        self.create_settings_page()
        self.create_logs_page()
        
        # Default Page
        self.show_dashboard()

    def setup_standard_ui(self):
        # Fallback simple UI if CustomTkinter is not found (prevents crash)
        self.sidebar_frame = tk.Frame(self.root, bg="#1e1e24", width=180)
        self.sidebar_frame.pack(side="left", fill="y")
        
        tk.Label(self.sidebar_frame, text="Outlook Sync", font=("Segoe UI", 14, "bold"), fg="white", bg="#1e1e24").pack(pady=20)
        
        tk.Button(self.sidebar_frame, text="Dashboard", bg="#2e2e38", fg="white", bd=0, relief="flat", height=2, command=self.show_dashboard).pack(fill="x", padx=10, pady=5)
        tk.Button(self.sidebar_frame, text="Configuration", bg="#2e2e38", fg="white", bd=0, relief="flat", height=2, command=self.show_settings).pack(fill="x", padx=10, pady=5)
        tk.Button(self.sidebar_frame, text="Console Logs", bg="#2e2e38", fg="white", bd=0, relief="flat", height=2, command=self.show_logs).pack(fill="x", padx=10, pady=5)
        
        self.content_frame = tk.Frame(self.root, bg="#282828")
        self.content_frame.pack(side="right", fill="both", expand=True)
        
        # Setup page views placeholder (simple warning)
        tk.Label(self.content_frame, text="Please run using 'run.bat' to install and run with beautiful CustomTkinter GUI.", fg="yellow", bg="#282828", font=("Segoe UI", 12)).pack(pady=100)
        
        # Add basic scrollable log screen
        self.log_textbox = tk.Text(self.content_frame, wrap="word", bg="#1e1e1e", fg="white")
        self.log_textbox.pack(fill="both", expand=True, padx=20, pady=20)
        
    def create_dashboard_page(self):
        self.dash_page = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        self.dash_page.grid_columnconfigure(0, weight=1)
        self.dash_page.grid_rowconfigure(2, weight=1)
        
        # Header Section
        header_frame = ctk.CTkFrame(self.dash_page, fg_color="#2b2b2b", corner_radius=8, height=70)
        header_frame.grid(row=0, column=0, padx=20, pady=20, sticky="ew")
        header_frame.grid_columnconfigure(0, weight=1)
        
        self.status_title = ctk.CTkLabel(
            header_frame, 
            text="Microsoft Outlook & SharePoint Automator", 
            font=ctk.CTkFont(size=20, weight="bold"),
            text_color="#ffffff"
        )
        self.status_title.grid(row=0, column=0, padx=20, pady=(15, 5), sticky="w")
        
        self.status_subtitle = ctk.CTkLabel(
            header_frame, 
            text="Idle • Ready to monitor 'HP Scan'", 
            font=ctk.CTkFont(size=12),
            text_color="#a0a0a0"
        )
        self.status_subtitle.grid(row=1, column=0, padx=20, pady=(0, 15), sticky="w")
        
        # Stats Cards Container
        stats_frame = ctk.CTkFrame(self.dash_page, fg_color="transparent")
        stats_frame.grid(row=1, column=0, padx=20, pady=(0, 20), sticky="ew")
        stats_frame.grid_columnconfigure((0, 1, 2), weight=1)
        
        # Card 1: Processed
        self.card_processed = ctk.CTkFrame(stats_frame, fg_color="#2b2b2b", height=100)
        self.card_processed.grid(row=0, column=0, padx=(0, 10), sticky="ew")
        self.lbl_proc_val = ctk.CTkLabel(self.card_processed, text="0", font=ctk.CTkFont(size=32, weight="bold"), text_color=self.accent_color)
        self.lbl_proc_val.pack(pady=(15, 5))
        ctk.CTkLabel(self.card_processed, text="Emails Filtered", font=ctk.CTkFont(size=12), text_color="#d0d0d0").pack(pady=(0, 15))
        
        # Card 2: Uploads
        self.card_uploaded = ctk.CTkFrame(stats_frame, fg_color="#2b2b2b", height=100)
        self.card_uploaded.grid(row=0, column=1, padx=5, sticky="ew")
        self.lbl_upload_val = ctk.CTkLabel(self.card_uploaded, text="0", font=ctk.CTkFont(size=32, weight="bold"), text_color="#107c41")
        self.lbl_upload_val.pack(pady=(15, 5))
        ctk.CTkLabel(self.card_uploaded, text="Attachments Uploaded", font=ctk.CTkFont(size=12), text_color="#d0d0d0").pack(pady=(0, 15))
        
        # Card 3: Duplicates
        self.card_skipped = ctk.CTkFrame(stats_frame, fg_color="#2b2b2b", height=100)
        self.card_skipped.grid(row=0, column=2, padx=(10, 0), sticky="ew")
        self.lbl_skip_val = ctk.CTkLabel(self.card_skipped, text="0", font=ctk.CTkFont(size=32, weight="bold"), text_color="#f2c12e")
        self.lbl_skip_val.pack(pady=(15, 5))
        ctk.CTkLabel(self.card_skipped, text="Duplicates Skipped", font=ctk.CTkFont(size=12), text_color="#d0d0d0").pack(pady=(0, 15))
        
        # Control Actions Frame
        control_frame = ctk.CTkFrame(self.dash_page, fg_color="#2b2b2b", corner_radius=8)
        control_frame.grid(row=2, column=0, padx=20, pady=(0, 20), sticky="nsew")
        control_frame.grid_columnconfigure(0, weight=1)
        control_frame.grid_rowconfigure(2, weight=1)
        
        ctk.CTkLabel(control_frame, text="Manual Execution Panel", font=ctk.CTkFont(size=14, weight="bold"), text_color="#ffffff").grid(row=0, column=0, padx=20, pady=(15, 10), sticky="w")
        
        # Run Button & Progress Area
        actions_inner = ctk.CTkFrame(control_frame, fg_color="transparent")
        actions_inner.grid(row=1, column=0, padx=20, pady=10, sticky="ew")
        actions_inner.grid_columnconfigure(0, weight=1)
        
        self.run_btn = ctk.CTkButton(
            actions_inner, 
            text="Sync Now", 
            font=ctk.CTkFont(size=16, weight="bold"),
            fg_color=self.accent_color,
            hover_color="#0066b3",
            height=45,
            command=self.run_sync_now
        )
        self.run_btn.grid(row=0, column=0, padx=(0, 20), sticky="ew")
        
        self.progress_bar = ctk.CTkProgressBar(control_frame, height=8, progress_color=self.accent_color)
        self.progress_bar.grid(row=3, column=0, padx=20, pady=(10, 20), sticky="ew")
        self.progress_bar.set(0.0)
        
        # Mini log box under control
        ctk.CTkLabel(control_frame, text="Quick Status Logs:", font=ctk.CTkFont(size=11), text_color="#a0a0a0").grid(row=4, column=0, padx=20, pady=0, sticky="w")
        self.mini_logs = ctk.CTkTextbox(control_frame, height=120, fg_color="#181818", text_color="#d0d0d0", font=ctk.CTkFont(family="Consolas", size=11))
        self.mini_logs.grid(row=5, column=0, padx=20, pady=(5, 20), sticky="ew")
        self.mini_logs.configure(state="disabled")

    def create_settings_page(self):
        self.settings_page = ctk.CTkScrollableFrame(self.content_frame, fg_color="transparent")
        self.settings_page.grid_columnconfigure(0, weight=1)
        
        # Sections: Outlook Folders, SharePoint, Auth Settings
        
        # 1. Outlook Settings
        outlook_sect = ctk.CTkFrame(self.settings_page, fg_color="#2b2b2b", corner_radius=8)
        outlook_sect.grid(row=0, column=0, padx=20, pady=10, sticky="ew")
        outlook_sect.grid_columnconfigure(1, weight=1)
        
        ctk.CTkLabel(outlook_sect, text="Outlook Mail Settings", font=ctk.CTkFont(size=14, weight="bold"), text_color="#ffffff").grid(row=0, column=0, columnspan=2, padx=15, pady=(15, 10), sticky="w")
        
        ctk.CTkLabel(outlook_sect, text="Target Folder Name:").grid(row=1, column=0, padx=15, pady=5, sticky="w")
        self.ent_outlook_folder = ctk.CTkEntry(outlook_sect)
        self.ent_outlook_folder.grid(row=1, column=1, padx=15, pady=5, sticky="ew")
        self.ent_outlook_folder.insert(0, self.config_data.get("outlook_folder"))
        
        ctk.CTkLabel(outlook_sect, text="Sender Email Filter:").grid(row=2, column=0, padx=15, pady=5, sticky="w")
        self.ent_sender = ctk.CTkEntry(outlook_sect)
        self.ent_sender.grid(row=2, column=1, padx=15, pady=5, sticky="ew")
        self.ent_sender.insert(0, self.config_data.get("sender_filter"))
        
        ctk.CTkLabel(outlook_sect, text="Subject Filter Words (comma separated):").grid(row=3, column=0, padx=15, pady=5, sticky="w")
        self.ent_subject = ctk.CTkEntry(outlook_sect)
        self.ent_subject.grid(row=3, column=1, padx=15, pady=5, sticky="ew")
        self.ent_subject.insert(0, self.config_data.get("subject_filter"))
        
        ctk.CTkLabel(outlook_sect, text="Outlook Source:").grid(row=4, column=0, padx=15, pady=5, sticky="w")
        self.var_outlook_source = tk.StringVar(value=self.config_data.get("outlook_source", "local"))
        
        outlook_source_frame = ctk.CTkFrame(outlook_sect, fg_color="transparent")
        outlook_source_frame.grid(row=4, column=1, padx=15, pady=5, sticky="w")
        
        self.rad_outlook_local = ctk.CTkRadioButton(outlook_source_frame, text="Classic Outlook App (Local COM)", variable=self.var_outlook_source, value="local")
        self.rad_outlook_local.pack(side="left", padx=(0, 20))
        
        self.rad_outlook_graph = ctk.CTkRadioButton(outlook_source_frame, text="Exchange Online (Graph API)", variable=self.var_outlook_source, value="graph")
        self.rad_outlook_graph.pack(side="left")
        
        # 2. SharePoint Settings
        sp_sect = ctk.CTkFrame(self.settings_page, fg_color="#2b2b2b", corner_radius=8)
        sp_sect.grid(row=1, column=0, padx=20, pady=10, sticky="ew")
        sp_sect.grid_columnconfigure(1, weight=1)
        
        ctk.CTkLabel(sp_sect, text="SharePoint & Storage Settings", font=ctk.CTkFont(size=14, weight="bold"), text_color="#ffffff").grid(row=0, column=0, columnspan=2, padx=15, pady=(15, 10), sticky="w")
        
        ctk.CTkLabel(sp_sect, text="Upload Method:").grid(row=1, column=0, padx=15, pady=5, sticky="w")
        self.var_sync_method = tk.StringVar(value=self.config_data.get("sync_method"))
        
        method_frame = ctk.CTkFrame(sp_sect, fg_color="transparent")
        method_frame.grid(row=1, column=1, padx=15, pady=5, sticky="w")
        
        self.rad_graph = ctk.CTkRadioButton(method_frame, text="Microsoft Graph Cloud API", variable=self.var_sync_method, value="graph", command=self.toggle_sync_method_fields)
        self.rad_graph.pack(side="left", padx=(0, 20))
        
        self.rad_local = ctk.CTkRadioButton(method_frame, text="OneDrive Local Synced Folder", variable=self.var_sync_method, value="local", command=self.toggle_sync_method_fields)
        self.rad_local.pack(side="left")
        
        ctk.CTkLabel(sp_sect, text="SharePoint Site URL:").grid(row=2, column=0, padx=15, pady=5, sticky="w")
        self.ent_sp_site = ctk.CTkEntry(sp_sect)
        self.ent_sp_site.grid(row=2, column=1, padx=15, pady=5, sticky="ew")
        self.ent_sp_site.insert(0, self.config_data.get("sharepoint_site"))
        
        ctk.CTkLabel(sp_sect, text="SharePoint Target Folder:").grid(row=3, column=0, padx=15, pady=5, sticky="w")
        self.ent_sp_folder = ctk.CTkEntry(sp_sect)
        self.ent_sp_folder.grid(row=3, column=1, padx=15, pady=5, sticky="ew")
        self.ent_sp_folder.insert(0, self.config_data.get("sharepoint_folder"))
        
        ctk.CTkLabel(sp_sect, text="OneDrive Folder Path:").grid(row=4, column=0, padx=15, pady=5, sticky="w")
        self.local_sync_frame = ctk.CTkFrame(sp_sect, fg_color="transparent")
        self.local_sync_frame.grid(row=4, column=1, padx=15, pady=5, sticky="ew")
        self.local_sync_frame.grid_columnconfigure(0, weight=1)
        
        self.ent_local_sync = ctk.CTkEntry(self.local_sync_frame)
        self.ent_local_sync.grid(row=0, column=0, sticky="ew", padx=(0, 10))
        self.ent_local_sync.insert(0, self.config_data.get("local_sync_path"))
        
        self.btn_browse = ctk.CTkButton(self.local_sync_frame, text="Browse...", width=80, command=self.browse_local_path)
        self.btn_browse.grid(row=0, column=1)
        
        # 3. Connection & API Auth Settings (Expandable or just visible)
        auth_sect = ctk.CTkFrame(self.settings_page, fg_color="#2b2b2b", corner_radius=8)
        auth_sect.grid(row=2, column=0, padx=20, pady=10, sticky="ew")
        auth_sect.grid_columnconfigure(1, weight=1)
        
        ctk.CTkLabel(auth_sect, text="Azure AD (Entra ID) API Credentials", font=ctk.CTkFont(size=14, weight="bold"), text_color="#ffffff").grid(row=0, column=0, columnspan=2, padx=15, pady=(15, 10), sticky="w")
        
        ctk.CTkLabel(auth_sect, text="Application (Client) ID:").grid(row=1, column=0, padx=15, pady=5, sticky="w")
        self.ent_client_id = ctk.CTkEntry(auth_sect)
        self.ent_client_id.grid(row=1, column=1, padx=15, pady=5, sticky="ew")
        self.ent_client_id.insert(0, self.config_data.get("client_id"))
        
        ctk.CTkLabel(auth_sect, text="Directory (Tenant) ID:").grid(row=2, column=0, padx=15, pady=5, sticky="w")
        self.ent_tenant_id = ctk.CTkEntry(auth_sect)
        self.ent_tenant_id.grid(row=2, column=1, padx=15, pady=5, sticky="ew")
        self.ent_tenant_id.insert(0, self.config_data.get("tenant_id"))
        
        # 4. Global Settings (Interval etc)
        global_sect = ctk.CTkFrame(self.settings_page, fg_color="#2b2b2b", corner_radius=8)
        global_sect.grid(row=3, column=0, padx=20, pady=10, sticky="ew")
        global_sect.grid_columnconfigure(1, weight=1)
        
        ctk.CTkLabel(global_sect, text="Sync Scheduling Settings", font=ctk.CTkFont(size=14, weight="bold"), text_color="#ffffff").grid(row=0, column=0, columnspan=2, padx=15, pady=(15, 10), sticky="w")
        
        ctk.CTkLabel(global_sect, text="Sync Interval (Minutes):").grid(row=1, column=0, padx=15, pady=5, sticky="w")
        self.ent_interval = ctk.CTkEntry(global_sect)
        self.ent_interval.grid(row=1, column=1, padx=15, pady=5, sticky="ew")
        self.ent_interval.insert(0, str(self.config_data.get("check_interval")))
        
        # Save Action Button
        self.btn_save = ctk.CTkButton(
            self.settings_page, 
            text="Save Settings", 
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="#107c41",
            hover_color="#0a5c30",
            height=40,
            command=self.save_settings
        )
        self.btn_save.grid(row=4, column=0, padx=20, pady=20, sticky="ew")
        
        # Run toggle styling immediately
        self.toggle_sync_method_fields()

    def create_logs_page(self):
        self.logs_page = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        self.logs_page.grid_columnconfigure(0, weight=1)
        self.logs_page.grid_rowconfigure(1, weight=1)
        
        # Title and Actions
        logs_header = ctk.CTkFrame(self.logs_page, fg_color="transparent")
        logs_header.grid(row=0, column=0, padx=20, pady=(10, 5), sticky="ew")
        logs_header.grid_columnconfigure(0, weight=1)
        
        ctk.CTkLabel(logs_header, text="System Log Console", font=ctk.CTkFont(size=16, weight="bold")).grid(row=0, column=0, sticky="w")
        
        btn_frame = ctk.CTkFrame(logs_header, fg_color="transparent")
        btn_frame.grid(row=0, column=1)
        
        self.btn_clear_logs = ctk.CTkButton(btn_frame, text="Clear Logs", width=80, fg_color="#b32e2e", hover_color="#8c2020", command=self.clear_logs)
        self.btn_clear_logs.pack(side="left", padx=5)
        
        self.btn_open_file = ctk.CTkButton(btn_frame, text="Open Log File", width=100, fg_color="#3e3e3e", hover_color="#2b2b2b", command=self.open_log_file)
        self.btn_open_file.pack(side="left", padx=5)
        
        # Console Log textbox
        self.log_console = ctk.CTkTextbox(self.logs_page, fg_color="#121212", text_color="#dfdfdf", font=ctk.CTkFont(family="Consolas", size=11))
        self.log_console.grid(row=1, column=0, padx=20, pady=(5, 20), sticky="nsew")
        self.log_console.configure(state="disabled")

    # View Management
    def show_dashboard(self):
        self.hide_all_pages()
        self.dash_page.grid(row=0, column=0, sticky="nsew")
        self.dash_btn.configure(fg_color="#282828", text_color="#ffffff")
        
    def show_settings(self):
        self.hide_all_pages()
        self.settings_page.grid(row=0, column=0, sticky="nsew")
        self.settings_btn.configure(fg_color="#282828", text_color="#ffffff")
        
    def show_logs(self):
        self.hide_all_pages()
        self.logs_page.grid(row=0, column=0, sticky="nsew")
        self.logs_btn.configure(fg_color="#282828", text_color="#ffffff")
        
    def hide_all_pages(self):
        self.dash_page.grid_forget()
        self.settings_page.grid_forget()
        self.logs_page.grid_forget()
        
        # Reset navigation button styling
        self.dash_btn.configure(fg_color="transparent", text_color="#d0d0d0")
        self.settings_btn.configure(fg_color="transparent", text_color="#d0d0d0")
        self.logs_btn.configure(fg_color="transparent", text_color="#d0d0d0")

    def toggle_sync_method_fields(self):
        method = self.var_sync_method.get()
        if method == "local":
            self.ent_sp_site.configure(state="disabled", fg_color="#202020", text_color="#606060")
            self.ent_local_sync.configure(state="normal", fg_color="#181818", text_color="#ffffff")
            self.btn_browse.configure(state="normal")
        else:
            self.ent_sp_site.configure(state="normal", fg_color="#181818", text_color="#ffffff")
            self.ent_local_sync.configure(state="disabled", fg_color="#202020", text_color="#606060")
            self.btn_browse.configure(state="disabled")
            
    def browse_local_path(self):
        path = filedialog.askdirectory(title="Select SharePoint OneDrive Synced Folder Root")
        if path:
            self.ent_local_sync.configure(state="normal")
            self.ent_local_sync.delete(0, tk.END)
            self.ent_local_sync.insert(0, path)
            self.toggle_sync_method_fields()
            
    def clear_logs(self):
        self.log_console.configure(state="normal")
        self.log_console.delete("1.0", tk.END)
        self.log_console.configure(state="disabled")
        
    def open_log_file(self):
        log_path = sync_engine.get_log_path()
        if os.path.exists(log_path):
            os.startfile(log_path)
        else:
            messagebox.showinfo("Information", "Log file does not exist yet. Run a sync session first.")
            
    def append_log(self, text, level=logging.INFO):
        # Format text with datetime
        timestamp = time.strftime("%H:%M:%S")
        entry = f"[{timestamp}] {text}\n"
        
        # Append to Main Log console
        if hasattr(self, 'log_console'):
            self.log_console.configure(state="normal")
            self.log_console.insert(tk.END, entry)
            self.log_console.see(tk.END)
            self.log_console.configure(state="disabled")
            
        # Append to Dashboard mini logs
        if hasattr(self, 'mini_logs'):
            self.mini_logs.configure(state="normal")
            self.mini_logs.insert(tk.END, entry)
            self.mini_logs.see(tk.END)
            self.mini_logs.configure(state="disabled")

    def save_settings(self):
        try:
            self.config_data["outlook_folder"] = self.ent_outlook_folder.get().strip()
            self.config_data["sender_filter"] = self.ent_sender.get().strip()
            self.config_data["subject_filter"] = self.ent_subject.get().strip()
            self.config_data["outlook_source"] = self.var_outlook_source.get()
            self.config_data["sync_method"] = self.var_sync_method.get()
            self.config_data["sharepoint_site"] = self.ent_sp_site.get().strip()
            self.config_data["sharepoint_folder"] = self.ent_sp_folder.get().strip()
            self.config_data["local_sync_path"] = self.ent_local_sync.get().strip()
            self.config_data["client_id"] = self.ent_client_id.get().strip()
            self.config_data["tenant_id"] = self.ent_tenant_id.get().strip()
            self.config_data["check_interval"] = int(self.ent_interval.get().strip())
            
            if config.save_config(self.config_data):
                messagebox.showinfo("Success", "Settings saved successfully.")
                self.append_log("Configuration saved successfully.", logging.INFO)
                # Restart scheduler with new interval
                self.restart_scheduler()
            else:
                messagebox.showerror("Error", "Could not save configuration file.")
        except ValueError:
            messagebox.showerror("Validation Error", "Sync interval must be a valid integer.")
        except Exception as e:
            messagebox.showerror("Error", f"Error saving settings: {e}")

    # Async Engine Executer
    def run_sync_now(self):
        if self.is_syncing:
            return
            
        self.is_syncing = True
        self.run_btn.configure(state="disabled", text="Running Sync...")
        self.progress_bar.set(0.0)
        self.progress_bar.configure(mode="indeterminate")
        self.progress_bar.start()
        
        self.status_subtitle.configure(text="Syncing • Running folder monitoring logic...", text_color="#0078d4")
        self.append_log("Starting manual synchronization...", logging.INFO)
        
        # Load latest saved configuration
        current_config = config.load_config()
        
        def run_thread():
            success, result = sync_engine.sync_run(current_config, self.log_progress_status)
            
            # Switch back to main UI thread to update GUI elements
            self.root.after(0, lambda: self.sync_completed(success, result))
            
        threading.Thread(target=run_thread, daemon=True).start()
        
    def log_progress_status(self, text, level):
        # Thread safe callback for logging progress
        self.root.after(0, lambda: self.append_log(text, level))
        
    def sync_completed(self, success, result):
        self.is_syncing = False
        self.progress_bar.stop()
        self.progress_bar.configure(mode="determinate")
        self.progress_bar.set(1.0 if success else 0.0)
        self.run_btn.configure(state="normal", text="Sync Now")
        
        if success:
            processed_emails, uploaded_files, skipped_files = result
            self.lbl_proc_val.configure(text=str(processed_emails))
            self.lbl_upload_val.configure(text=str(uploaded_files))
            self.lbl_skip_val.configure(text=str(skipped_files))
            
            status_msg = f"Idle • Last sync completed successfully. ({uploaded_files} uploaded, {skipped_files} skipped)"
            self.status_subtitle.configure(text=status_msg, text_color="#107c41")
            self.append_log("Synchronization completed successfully.", logging.INFO)
        else:
            self.status_subtitle.configure(text=f"Error • Sync failed: {result}", text_color="#b32e2e")
            self.append_log(f"Sync failed with error: {result}", logging.ERROR)
            messagebox.showerror("Sync Failed", f"Synchronization encountered an error:\n{result}")

    # Background Scheduler Logic
    def start_scheduler(self):
        if self.scheduler_running:
            return
            
        self.scheduler_running = True
        self.scheduler_stop_event = threading.Event()
        if hasattr(self, 'sched_badge'):
            self.sched_badge.configure(text="Scheduler: ACTIVE", text_color="#107c41")
            
        def scheduler_loop(stop_event):
            # Initial sleep before first run or run immediately?
            # Let's run after the interval has elapsed
            self.append_log(f"Background monitoring scheduled to run every {self.config_data.get('check_interval')} minutes.", logging.INFO)
            
            while not stop_event.is_set():
                interval_minutes = self.config_data.get("check_interval", 15)
                # Sleep in increments of 1 second to allow fast cancellation
                for _ in range(interval_minutes * 60):
                    if stop_event.is_set():
                        break
                    time.sleep(1)
                    
                if stop_event.is_set():
                    break
                    
                # Run sync in background (if not already syncing)
                if not self.is_syncing:
                    self.append_log("Scheduled trigger: Starting automatic synchronization...", logging.INFO)
                    self.root.after(0, self.run_sync_now_silent)
                    
        self.scheduler_thread = threading.Thread(target=scheduler_loop, args=(self.scheduler_stop_event,), daemon=True)
        self.scheduler_thread.start()
        
    def stop_scheduler(self):
        self.scheduler_running = False
        if self.scheduler_stop_event is not None:
            self.scheduler_stop_event.set()
        if hasattr(self, 'sched_badge'):
            self.sched_badge.configure(text="Scheduler: STOPPED", text_color="#b32e2e")
        self.append_log("Background scheduler stopped.", logging.INFO)
        
    def restart_scheduler(self):
        self.stop_scheduler()
        # Wait a moment for thread to exit
        time.sleep(0.5)
        self.config_data = config.load_config()
        self.start_scheduler()
        
    def run_sync_now_silent(self):
        # Triggered by scheduler, updates stats and status quietly
        if self.is_syncing:
            return
            
        self.is_syncing = True
        self.progress_bar.set(0.0)
        self.progress_bar.configure(mode="indeterminate")
        self.progress_bar.start()
        self.status_subtitle.configure(text="Syncing (Auto) • Running background monitor...", text_color="#0078d4")
        
        current_config = config.load_config()
        
        def run_thread():
            success, result = sync_engine.sync_run(current_config, self.log_progress_status)
            self.root.after(0, lambda: self.sync_completed_silent(success, result))
            
        threading.Thread(target=run_thread, daemon=True).start()
        
    def sync_completed_silent(self, success, result):
        self.is_syncing = False
        self.progress_bar.stop()
        self.progress_bar.configure(mode="determinate")
        self.progress_bar.set(1.0 if success else 0.0)
        
        if success:
            processed_emails, uploaded_files, skipped_files = result
            # Sum up stats
            try:
                cur_proc = int(self.lbl_proc_val.cget("text")) + processed_emails
                cur_upload = int(self.lbl_upload_val.cget("text")) + uploaded_files
                cur_skip = int(self.lbl_skip_val.cget("text")) + skipped_files
                
                self.lbl_proc_val.configure(text=str(cur_proc))
                self.lbl_upload_val.configure(text=str(cur_upload))
                self.lbl_skip_val.configure(text=str(cur_skip))
            except Exception:
                pass
                
            status_msg = f"Idle • Auto sync finished. ({uploaded_files} uploaded, {skipped_files} skipped)"
            self.status_subtitle.configure(text=status_msg, text_color="#107c41")
            self.append_log("Auto sync completed successfully.", logging.INFO)
        else:
            self.status_subtitle.configure(text=f"Error • Auto sync failed", text_color="#b32e2e")
            self.append_log(f"Auto sync failed with error: {result}", logging.ERROR)

def main():
    # Set DPI awareness for high-DPI displays (makes text crisp on Win 11)
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
        
    if ctk_available:
        root = ctk.CTk()
    else:
        root = tk.Tk()
        
    app = OutlookSyncApp(root)
    
    # Handle closing clean
    def on_closing():
        app.stop_scheduler()
        root.destroy()
        
    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.mainloop()

if __name__ == "__main__":
    main()
