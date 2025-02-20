############################################################
# Win11 Debloat GUI
# Version 1.0
#
# A simple GUI for the Win11Debloat PowerShell script.
# Allows users to select debloat options and run the script.
#
# Authors: Alp Orak, M.C.Aksoy
# Date: 2025
############################################################

# Import required modules
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
import subprocess
from sys import argv
from sys import executable
from ctypes import windll
import threading
from json import dump, load
import os

class Win11DebloatUI:
    ''' Main application class for the Win11 Debloat GUI '''
    def __init__(self, root):
        self.root = root
        self.root.title("Win11 Debloat GUI")
        self.root.geometry("1000x700")
        self.process = None
        self.running = False
        self.settings_file = "debloat_settings.json"
        self.admin_mode = self.is_admin()
        
        self.create_widgets()
        self.create_menu()
        self.load_settings()
        self.create_parameter_panel()
        self.update_ui_state()

    def is_admin(self):
        try:
            return windll.shell32.IsUserAnAdmin()
        except:
            return False

    def restart_as_admin(self):
        try:
            # Launch new admin instance
            windll.shell32.ShellExecuteW(
                None,
                "runas",
                executable,
                " ".join(argv),
                None,
                1
            )
        finally:
            # Ensure original instance always closes
            self.root.destroy()
            os._exit(0)


    def update_ui_state(self):
        # Update admin button visibility
        if self.admin_mode:
            self.get_admin.pack_forget()
            self.start_btn.pack(side=tk.LEFT, padx=5)
        else:
            self.get_admin.pack(side=tk.LEFT, padx=5)
            self.start_btn.pack_forget()
        
        # Update cancel button visibility
        self.cancel_btn.pack_forget()

    def create_menu(self):
        menubar = tk.Menu(self.root)
        
        # File menu
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Save Settings", command=self.save_settings)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)
        menubar.add_cascade(label="File", menu=file_menu)

        # Tools menu (only show in admin mode)
        if self.admin_mode:
            tools_menu = tk.Menu(menubar, tearoff=0)
            tools_menu.add_command(label="App Configurator", command=self.show_app_configurator)
            menubar.add_cascade(label="Tools", menu=tools_menu)

        
        # Help menu
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="About", command=lambda: messagebox.showinfo("About", "Win11 Debloat GUI\nVersion 1.0"))
        help_menu.add_command(label="Check for Updates", command=lambda: messagebox.showinfo("Updates", "No updates available.")) # TODO: Implement update check
        help_menu.add_command(label="Help", command=lambda: messagebox.showinfo("Help", "Select the options you want to apply and click 'Start Debloat' to begin the process."))
        menubar.add_cascade(label="Help", menu=help_menu)

        self.root.config(menu=menubar)

    def create_parameter_panel(self):
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill='both', expand=True, padx=10, pady=5)

        # System Tweaks Tab
        system_frame = ttk.Frame(notebook)
        self.create_section(system_frame, "System Tweaks", [
            ('DisableTelemetry', 'Disable Telemetry'),
            ('DisableBing', 'Disable Bing Search/Cortana'),
            ('DisableWidgets', 'Disable Widgets'),
            ('DisableCopilot', 'Disable Copilot'),
            ('DisableRecall', 'Disable Recall Snapshots')
        ])
        notebook.add(system_frame, text="System Tweaks")

        # Apps Tab
        apps_frame = ttk.Frame(notebook)
        self.create_section(apps_frame, "Application Management", [
            ('RemoveApps', 'Remove Default Apps'),
            ('RemoveGamingApps', 'Remove Gaming Apps'),
            ('RemoveCommApps', 'Remove Communication Apps'),
            ('ForceRemoveEdge', 'Force Remove Edge Browser')
        ])
        notebook.add(apps_frame, text="Applications")

        # UI Tweaks Tab
        ui_frame = ttk.Frame(notebook)
        self.create_section(ui_frame, "UI Customization", [
            ('TaskbarAlignLeft', 'Align Taskbar Left'),
            ('RevertContextMenu', 'Classic Context Menu'),
            ('ShowHiddenFolders', 'Show Hidden Files'),
            ('ShowKnownFileExt', 'Show File Extensions')
        ])
        notebook.add(ui_frame, text="UI Tweaks")

    def create_section(self, parent, title, options):
        frame = ttk.LabelFrame(parent, text=title)
        frame.pack(fill='both', expand=True, padx=5, pady=5)
        
        for var_name, label in options:
            cb = ttk.Checkbutton(
                frame,
                text=label,
                variable=self.settings[var_name],
                onvalue=True,
                offvalue=False
            )
            cb.pack(anchor=tk.W, padx=5, pady=2)

    def create_widgets(self):
        # Terminal output
        self.terminal = scrolledtext.ScrolledText(self.root, wrap=tk.WORD)
        self.terminal.pack(fill='both', expand=True, padx=10, pady=5)

        # Progress bar
        self.progress = ttk.Progressbar(self.root, orient=tk.HORIZONTAL, mode='determinate')
        self.progress.pack(fill=tk.X, padx=10, pady=5)

        # Control buttons
        button_frame = ttk.Frame(self.root)
        button_frame.pack(pady=10)

        self.start_btn = ttk.Button(
            button_frame, 
            text="Start Debloat", 
            command=self.start_debloat
        )

        self.get_admin = ttk.Button(
            button_frame,
            text="Restart as Admin",
            command=self.restart_as_admin
        )

        self.cancel_btn = ttk.Button(
            button_frame, 
            text="Cancel", 
            command=self.cancel_operation
        )

    def load_settings(self):
        default_settings = {
            'DisableTelemetry': tk.BooleanVar(value=True),
            'DisableBing': tk.BooleanVar(value=True),
            'RemoveApps': tk.BooleanVar(value=True),
            'RemoveGamingApps': tk.BooleanVar(value=False),
            'RemoveCommApps': tk.BooleanVar(value=False),
            'ForceRemoveEdge': tk.BooleanVar(value=False),
            'TaskbarAlignLeft': tk.BooleanVar(value=True),
            'RevertContextMenu': tk.BooleanVar(value=True),
            'ShowHiddenFolders': tk.BooleanVar(value=True),
            'ShowKnownFileExt': tk.BooleanVar(value=True),
            'DisableWidgets': tk.BooleanVar(value=True),
            'DisableCopilot': tk.BooleanVar(value=True),
            'DisableRecall': tk.BooleanVar(value=True)
        }
        
        try:
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r') as f:
                    self.log("Found settings file, loading...")
                    saved_settings = load(f)
                    self.settings = {
                        key: tk.BooleanVar(value=value) 
                        for key, value in saved_settings.items()
                    }
            else:
                self.settings = default_settings
        except:
            self.settings = default_settings

    def save_settings(self):
        try:
            with open(self.settings_file, 'w') as f:
                settings_to_save = {key: var.get() for key, var in self.settings.items()}
                dump(settings_to_save, f)
            self.log("Settings saved successfully.")
        except Exception as e:
            self.log(f"Error saving settings: {str(e)}")

    def start_debloat(self):
        if self.running:
            messagebox.showwarning("Already Running", "A debloat operation is already in progress!")
            return

        self.running = True
        self.start_btn.config(state=tk.DISABLED)
        self.cancel_btn.pack(side=tk.LEFT, padx=5)
        self.terminal.delete(1.0, tk.END)
        self.progress['value'] = 0

        params = [key for key, var in self.settings.items() if var.get()]
        threading.Thread(target=self.run_debloat_script, args=(params,), daemon=True).start()

    def run_debloat_script(self, params):
        try:
            self.log("Starting debloat process...")
            
            command = [
                "powershell.exe",
                "-ExecutionPolicy", "Bypass",
                "-File", "Win11Debloat\Win11Debloat.ps1"
            ] + [f"-{param}" for param in params]

            self.process = subprocess.Popen(
                command,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                creationflags=subprocess.CREATE_NO_WINDOW
            )

            while True:
                output = self.process.stdout.readline()
                if output == '' and self.process.poll() is not None:
                    break
                if output:
                    self.log(output.strip())
                
                self.root.update_idletasks()

            return_code = self.process.poll()
            if return_code == 0:
                self.log("Debloat process completed successfully!")
                messagebox.showinfo("Success", "Debloat operation completed successfully!")
            else:
                self.log(f"Debloat process failed with return code {return_code}")
                messagebox.showerror("Error", "Debloat operation encountered errors!")

        except Exception as e:
            self.log(f"Error: {str(e)}")
            messagebox.showerror("Critical Error", str(e))
        finally:
            self.running = False
            self.start_btn.config(state=tk.NORMAL)
            self.progress['value'] = 100
            self.cancel_btn.pack_forget()

    def show_app_configurator(self):
        if not self.admin_mode:
            messagebox.showerror("Permission Denied", "This feature requires admin privileges!")
            return
        try:
            subprocess.Popen([
                "powershell.exe",
                "-ExecutionPolicy", "Bypass",
                "-File", "Win11Debloat\Win11Debloat.ps1",
                "-RunAppConfigurator"
            ], creationflags=subprocess.CREATE_NO_WINDOW)
        except Exception as e:
            self.log(f"Error launching app configurator: {str(e)}")

    def log(self, message):
        self.terminal.insert(tk.END, message + "\n")
        self.terminal.see(tk.END)
        self.root.update_idletasks()

    def cancel_operation(self):
        if self.process and self.running:
            self.process.terminate()
            self.log("Operation cancelled by user")
            self.running = False
            self.start_btn.config(state=tk.NORMAL)
            self.progress['value'] = 0
            self.cancel_btn.pack_forget()

if __name__ == "__main__":
    root = tk.Tk()
    app = Win11DebloatUI(root)
    root.mainloop()