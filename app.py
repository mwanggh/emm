import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import threading
import subprocess
import ctypes
import atexit
import sys
import os


try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    pass


def get_resource_path(relative_path):
    """获取资源文件的路径，无论是开发环境还是打包后"""
    if hasattr(sys, '_MEIPASS'):  # PyInstaller 打包后的临时路径
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)


def get_running_processes():
    # Retrieve a sorted list of running process names.
    command = 'Get-Process | Select-Object -Property Name -Unique | Sort-Object Name | ForEach-Object { $_.Name }'
    result = subprocess.run(
        ["powershell", "-Command", command], 
        capture_output=True, 
        text=True,
        creationflags=subprocess.CREATE_NO_WINDOW
    )
    return result.stdout.strip().split("\n")


def generate_monitor_script(programs):
    # Generate a PowerShell script to monitor and adjust process priorities.
    programs = ", ".join([f'"{p}"' for p in programs])
    return f"""
$programs = @({programs})

while ($true) {{
    foreach ($program in $programs) {{
        $processes = Get-Process -Name $program -ErrorAction SilentlyContinue
        foreach ($process in $processes) {{
            if ($process.PriorityClass -eq [System.Diagnostics.ProcessPriorityClass]::Idle) {{
                $process.PriorityClass = [System.Diagnostics.ProcessPriorityClass]::Normal
                Write-Output "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')|$($program)|Changed priority: Idle to Normal."
            }}
        }}
    }}
    Start-Sleep -Milliseconds 200
}}
"""


class EfficiencyApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Disable Efficiency Mode")
        self.root.iconbitmap(get_resource_path("res/rocket.ico"))
        self.root.geometry("1200x800")

        self.closed_programs = []
        self.monitor_process = None

        self.load_config()

        # Configure styles
        self.style = ttk.Style()
        self.style.configure("Treeview.Heading", anchor="w")  # Align Treeview headers
        self.style.configure("Treeview", rowheight=30)  # Set Treeview row height
        self.style.configure("Green.TButton", foreground="darkgreen")

        # Dropdown label
        self.label = ttk.Label(root, text="Select or enter a program: ")
        self.label.grid(row=0, column=0, sticky="e", padx=10)

        # Dropdown for running processes
        self.program_var = tk.StringVar()
        self.program_dropdown = ttk.Combobox(
            root, textvariable=self.program_var, values=get_running_processes(), state="normal"
        )
        self.program_dropdown.grid(row=0, column=1, sticky="w", padx=10)
        self.program_dropdown.bind("<Button-1>", self.on_dropdown_click)

        # Button to close efficiency mode
        self.close_button = ttk.Button(root, text="Disable Efficiency Mode", command=self.close_efficiency_mode)
        self.close_button.grid(row=1, column=1, sticky="w", padx=10, pady=10)

        # Listbox for selected programs
        self.programs_listbox = tk.Listbox(root, selectmode=tk.SINGLE)
        self.programs_listbox.grid(row=3, column=0, sticky="nsew", padx=10)

        # Treeview for logs
        self.log_tree = ttk.Treeview(root, columns=("Time", "Software", "Action"), show="headings")
        self.log_tree.grid(row=3, column=1, rowspan=2, sticky="nsew", padx=10)

        # Treeview headers and columns
        self.log_tree.heading("Time", text="Time")
        self.log_tree.heading("Software", text="Software")
        self.log_tree.heading("Action", text="Action")
        self.log_tree.column("Time", width=200, anchor="w")
        self.log_tree.column("Software", width=100, anchor="w")
        self.log_tree.column("Action", width=400, anchor="w")

        # Grid layout weights
        self.root.grid_rowconfigure(3, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_columnconfigure(1, weight=3)

        # Button to delete selected programs
        self.delete_button = ttk.Button(root, text="Delete", command=self.delete_selected_program)
        self.delete_button.grid(row=4, column=0, sticky="nsew", padx=10)

        # Button to start/stop monitoring
        self.monitor_button = ttk.Button(root, text="Start Monitoring", command=self.toggle_monitor)
        self.monitor_button.grid(row=5, column=0, sticky="w", padx=10, pady=10)

        # Update Listbox with initial data
        self.update_programs_listbox()

        # Start monitoring automatically
        self.toggle_monitor()

        # Register cleanup function to ensure the shell process is terminated on exit
        atexit.register(self.cleanup)

    def on_dropdown_click(self, event):
        # Update dropdown menu with running processes
        running_processes = get_running_processes()
        self.program_dropdown["values"] = running_processes  

    def close_efficiency_mode(self):
        # Add selected program to closed list if not already added
        selected_program = self.program_var.get()
        if not selected_program or selected_program in self.closed_programs:
            return
        self.closed_programs.append(selected_program)
        self.save_config()
        self.update_programs_listbox()
    
    def update_programs_listbox(self):
        # Refresh Listbox with closed programs
        self.programs_listbox.delete(0, tk.END)
        for program in self.closed_programs:
            self.programs_listbox.insert(tk.END, program)
        
        if self.monitor_process:
            self.stop_monitor()
            self.start_monitor()
    
    def delete_selected_program(self):
        # Remove selected program from closed list
        selected_index = self.programs_listbox.curselection()
        if selected_index:
            selected_program = self.programs_listbox.get(selected_index)
            self.closed_programs.remove(selected_program)
            self.save_config()
            self.update_programs_listbox()
    
    def start_monitor(self):
        try:
            # Start monitoring process
            script_content = generate_monitor_script(self.closed_programs)
            self.monitor_process = subprocess.Popen(
                ["powershell", "-Command", script_content],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            threading.Thread(target=self.read_logs, daemon=True).start()
            self.monitor_button.config(text="Monitoring")
            self.monitor_button.configure(style="Green.TButton")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to start monitoring: {e}")
            self.monitor_button.config(text="Start Monitoring")
            self.monitor_button.configure(style="TButton")

    def stop_monitor(self):
        # Stop monitoring process
        if self.monitor_process:
            self.monitor_process.terminate()
            self.monitor_process = None
            self.monitor_button.config(text="Start Monitoring")
            self.monitor_button.configure(style="TButton")
    
    def read_logs(self):
        # Read and parse logs from the monitoring process
        while self.monitor_process and self.monitor_process.stdout:
            log_line = self.monitor_process.stdout.readline()
            if log_line:
                parts = log_line.strip().split("|")
                if len(parts) == 3:
                    timestamp, program, action = parts
                    self.log_tree.insert("", "0", values=(timestamp, program, action))
    
    def toggle_monitor(self):
        # Toggle between start and stop monitoring
        if self.monitor_process:
            self.stop_monitor()
        else:
            self.start_monitor()
    
    def load_config(self):
        # Load closed programs from configuration file
        try:
            with open("monitor_sc.txt", "r") as config_file:
                lines = config_file.readlines()
                for line in lines:
                    parts = line.strip().split("|")
                    if len(parts) == 2 and parts[1] == "CLOSE":
                        self.closed_programs.append(parts[0])
        except FileNotFoundError:
            pass
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load configuration: {e}")

    def save_config(self):
        # Save closed programs to configuration file
        try:
            with open("monitor_sc.txt", "w") as config_file:
                config_file.write("\n".join([f"{program}|CLOSE" for program in self.closed_programs]))
        except Exception as e:
            ttk.messagebox.showerror("Error", f"Failed to save configuration: {e}")
    
    def cleanup(self):
        # Terminate the PowerShell process if it's running
        if self.monitor_process:
            self.monitor_process.terminate()
            self.monitor_process = None


if __name__ == "__main__":
    root = tk.Tk()
    app = EfficiencyApp(root)
    root.mainloop()

# pyinstaller --onefile --noconsole --icon=res/rocket.ico app.py; Copy-Item .\dist\app.exe .\ -Force
# pyinstaller --onefile --noconsole --icon=res/rocket.ico --add-data "res/rocket.ico;res" app.py; Copy-Item .\dist\app.exe .\ -Force