import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import socket
import json
import threading
from cryptography.fernet import Fernet
import customtkinter as ctk
import time


class MCCClient(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Initialize connection settings
        self.server_host = 'localhost'
        self.server_port = 5000
        self.socket = None
        self.connected = False
        self.encryption_key = None
        self.cipher_suite = None

        # Configure the window
        self.title("Multi Computers Control")
        self.geometry("1200x800")

        # Set the theme
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        self.create_gui()
        self.initialize_monitoring()
        self.connect_to_server()

    def create_gui(self):
        """Create the main GUI interface"""
        # Create main container
        self.main_container = ctk.CTkFrame(self)
        self.main_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Create sidebar for computer list
        self.sidebar = ctk.CTkFrame(self.main_container, width=200)
        self.sidebar.pack(side=tk.LEFT, fill=tk.Y, padx=5, pady=5)

        # Create computer list
        self.computer_list = ctk.CTkTextbox(self.sidebar, width=200)
        self.computer_list.pack(fill=tk.BOTH, expand=True)

        # Create main content area with tabs
        self.notebook = ttk.Notebook(self.main_container)
        self.notebook.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Create tabs
        self.create_monitoring_tab()
        self.create_software_tab()
        self.create_power_tab()
        self.create_file_transfer_tab()
        self.create_command_tab()
        self.create_network_tab()

    def create_monitoring_tab(self):
        """Create the monitoring tab"""
        monitoring_frame = ctk.CTkFrame(self.notebook)
        self.notebook.add(monitoring_frame, text="Monitoring")

        # CPU Usage
        cpu_frame = ctk.CTkFrame(monitoring_frame)
        cpu_frame.pack(fill=tk.X, padx=5, pady=5)

        ctk.CTkLabel(cpu_frame, text="CPU Usage:").pack(side=tk.LEFT)
        self.cpu_progress = ctk.CTkProgressBar(cpu_frame)
        self.cpu_progress.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.cpu_progress.set(0)

        # Memory Usage
        mem_frame = ctk.CTkFrame(monitoring_frame)
        mem_frame.pack(fill=tk.X, padx=5, pady=5)

        ctk.CTkLabel(mem_frame, text="Memory Usage:").pack(side=tk.LEFT)
        self.mem_progress = ctk.CTkProgressBar(mem_frame)
        self.mem_progress.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.mem_progress.set(0)

        # Disk Usage
        self.disk_frame = ctk.CTkFrame(monitoring_frame)
        self.disk_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # System Information
        self.system_info = ctk.CTkTextbox(monitoring_frame, height=100)
        self.system_info.pack(fill=tk.X, padx=5, pady=5)

    def create_software_tab(self):
        """Create the software management tab"""
        software_frame = ctk.CTkFrame(self.notebook)
        self.notebook.add(software_frame, text="Software")

        # Software list
        self.software_tree = ttk.Treeview(software_frame, columns=("Name", "Version"), show="headings")
        self.software_tree.heading("Name", text="Name")
        self.software_tree.heading("Version", text="Version")
        self.software_tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Control buttons
        button_frame = ctk.CTkFrame(software_frame)
        button_frame.pack(fill=tk.X, padx=5, pady=5)

        ctk.CTkButton(button_frame, text="Refresh", command=self.refresh_software_list).pack(side=tk.LEFT, padx=5)
        ctk.CTkButton(button_frame, text="Install", command=self.install_software).pack(side=tk.LEFT, padx=5)
        ctk.CTkButton(button_frame, text="Uninstall", command=self.uninstall_software).pack(side=tk.LEFT, padx=5)

    def create_power_tab(self):
        """Create the power management tab"""
        power_frame = ctk.CTkFrame(self.notebook)
        self.notebook.add(power_frame, text="Power")

        # Power management buttons
        button_frame = ctk.CTkFrame(power_frame)
        button_frame.pack(expand=True)

        ctk.CTkButton(button_frame, text="Shutdown", command=lambda: self.power_action("shutdown")).pack(pady=5)
        ctk.CTkButton(button_frame, text="Restart", command=lambda: self.power_action("restart")).pack(pady=5)
        ctk.CTkButton(button_frame, text="Sleep", command=lambda: self.power_action("sleep")).pack(pady=5)

        # Schedule frame
        schedule_frame = ctk.CTkFrame(power_frame)
        schedule_frame.pack(expand=True, pady=20)

        ctk.CTkLabel(schedule_frame, text="Schedule Power Action").pack()

        self.schedule_time = ctk.CTkEntry(schedule_frame, placeholder_text="HH:MM")
        self.schedule_time.pack(pady=5)

        ctk.CTkButton(schedule_frame, text="Schedule", command=self.schedule_power_action).pack(pady=5)

    def create_file_transfer_tab(self):
        """Create the file transfer tab"""
        file_frame = ctk.CTkFrame(self.notebook)
        self.notebook.add(file_frame, text="File Transfer")

        # Source selection
        source_frame = ctk.CTkFrame(file_frame)
        source_frame.pack(fill=tk.X, padx=5, pady=5)

        ctk.CTkButton(source_frame, text="Select Files", command=self.select_files).pack(side=tk.LEFT, padx=5)
        self.source_path = ctk.CTkEntry(source_frame)
        self.source_path.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

        # Destination selection
        dest_frame = ctk.CTkFrame(file_frame)
        dest_frame.pack(fill=tk.X, padx=5, pady=5)

        ctk.CTkButton(dest_frame, text="Select Destination", command=self.select_destination).pack(side=tk.LEFT, padx=5)
        self.dest_path = ctk.CTkEntry(dest_frame)
        self.dest_path.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

        # Transfer progress
        self.transfer_progress = ctk.CTkProgressBar(file_frame)
        self.transfer_progress.pack(fill=tk.X, padx=5, pady=5)
        self.transfer_progress.set(0)

        # Transfer button
        ctk.CTkButton(file_frame, text="Transfer", command=self.transfer_files).pack(pady=5)

        # Transfer log
        self.transfer_log = ctk.CTkTextbox(file_frame)
        self.transfer_log.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

    def create_command_tab(self):
        """Create the command execution tab"""
        command_frame = ctk.CTkFrame(self.notebook)
        self.notebook.add(command_frame, text="Commands")

        # Command input
        input_frame = ctk.CTkFrame(command_frame)
        input_frame.pack(fill=tk.X, padx=5, pady=5)

        self.command_input = ctk.CTkEntry(input_frame, placeholder_text="Enter command...")
        self.command_input.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

        ctk.CTkButton(input_frame, text="Execute", command=self.execute_command).pack(side=tk.LEFT)

        # Command output
        self.command_output = ctk.CTkTextbox(command_frame)
        self.command_output.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Saved commands
        saved_frame = ctk.CTkFrame(command_frame)
        saved_frame.pack(fill=tk.X, padx=5, pady=5)

        ctk.CTkLabel(saved_frame, text="Saved Commands:").pack(side=tk.LEFT)
        self.saved_commands = ctk.CTkComboBox(saved_frame, values=["dir", "ipconfig", "systeminfo"])
        self.saved_commands.pack(side=tk.LEFT, padx=5)

        ctk.CTkButton(saved_frame, text="Run", command=self.run_saved_command).pack(side=tk.LEFT)

    def create_network_tab(self):
        """Create the network monitoring tab"""
        network_frame = ctk.CTkFrame(self.notebook)
        self.notebook.add(network_frame, text="Network")

        # Network statistics
        stats_frame = ctk.CTkFrame(network_frame)
        stats_frame.pack(fill=tk.X, padx=5, pady=5)

        self.network_stats = ctk.CTkTextbox(stats_frame, height=100)
        self.network_stats.pack(fill=tk.X, expand=True)

        # Active connections
        connections_frame = ctk.CTkFrame(network_frame)
        connections_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        self.connections_tree = ttk.Treeview(connections_frame,
                                             columns=("Local", "Remote", "Status"), show="headings")
        self.connections_tree.heading("Local", text="Local Address")
        self.connections_tree.heading("Remote", text="Remote Address")
        self.connections_tree.heading("Status", text="Status")
        self.connections_tree.pack(fill=tk.BOTH, expand=True)

    def connect_to_server(self):
        """Establish connection to the server"""
        try:
            self.socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            self.socket.connect((self.server_host, self.server_port))

            # Wait for encryption key from server
            self.encryption_key = self.socket.recv(1024)
            if not self.encryption_key:
                raise Exception("Did not receive encryption key from server")

            self.cipher_suite = Fernet(self.encryption_key)
            self.connected = True

            # Start receive thread
            threading.Thread(target=self.receive_data, daemon=True).start()

            messagebox.showinfo("Connection", "Successfully connected to server")

        except Exception as e:
            self.connected = False
            self.cipher_suite = None
            messagebox.showerror("Connection Error", f"Failed to connect: {str(e)}")

    def send_command(self, command_type, data):
        """Send encrypted command to server"""
        if not self.connected:
            return {'status': 'error', 'message': 'Not connected to server'}

        command = {
            'type': command_type,
            'data': data
        }

        encrypted_data = self.cipher_suite.encrypt(json.dumps(command).encode())
        self.socket.send(encrypted_data)

        # Wait for response
        encrypted_response = self.socket.recv(4096)
        response = json.loads(self.cipher_suite.decrypt(encrypted_response).decode())
        return response

    def receive_data(self):
        """Handle incoming data from server"""
        while self.connected:
            try:
                encrypted_data = self.socket.recv(4096)
                if not encrypted_data:
                    break

                data = json.loads(self.cipher_suite.decrypt(encrypted_data).decode())
                self.handle_server_data(data)

            except Exception as e:
                print(f"Error receiving data: {str(e)}")
                break

        self.connected = False
        messagebox.showwarning("Connection Lost", "Lost connection to server")

    def handle_server_data(self, data):
        """Handle different types of data from server"""
        data_type = data.get('type')
        if data_type == 'system_info':
            self.update_system_info(data['data'])
        elif data_type == 'hardware_monitor':
            self.update_hardware_info(data['data'])
        elif data_type == 'network_monitor':
            self.update_network_info(data['data'])

    def initialize_monitoring(self):
        """Initialize monitoring thread"""
        self.monitoring_thread = threading.Thread(target=self.monitor_resources, daemon=True)
        self.monitoring_thread.start()

    def monitor_resources(self):
        """Monitor system resources with improved error handling"""
        while True:
            if self.connected and self.cipher_suite:
                try:
                    # Get hardware info
                    response = self.send_command('hardware_monitor', {})
                    if response and response.get('status') == 'success':
                        self.after(0, self.update_hardware_info, response['data'])

                    # Get system info for computer list
                    response = self.send_command('system_info', {})
                    if response and response.get('status') == 'success':
                        self.after(0, self.update_computer_list, [response['data']])

                except Exception as e:
                    print(f"Monitoring error: {str(e)}")

            time.sleep(1)  # Update every second

    def update_hardware_info(self, data):
        """Update hardware monitoring displays"""
        try:
            # Update CPU usage
            cpu_percent = data['cpu_percent']
            self.cpu_progress.configure(mode="determinate")  # Ensure progress bar is in determinate mode
            self.cpu_progress.set(cpu_percent / 100.0)  # Convert percentage to 0-1 range

            # Update memory usage
            memory_data = data['memory_usage']
            memory_percent = memory_data['percent']
            self.mem_progress.configure(mode="determinate")
            self.mem_progress.set(memory_percent / 100.0)

            # Update disk usage display
            for widget in self.disk_frame.winfo_children():
                widget.destroy()

            for mount, usage in data['disk_usage'].items():
                frame = ctk.CTkFrame(self.disk_frame)
                frame.pack(fill=tk.X, padx=5, pady=2)

                # Add drive label
                label = ctk.CTkLabel(frame, text=f"{mount}")
                label.pack(side=tk.LEFT, padx=5)

                # Add progress bar
                progress = ctk.CTkProgressBar(frame)
                progress.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
                progress.configure(mode="determinate")
                progress.set(usage['percent'] / 100.0)

                # Add percentage label
                percent_label = ctk.CTkLabel(frame, text=f"{usage['percent']:.1f}%")
                percent_label.pack(side=tk.RIGHT, padx=5)

            self.update()  # Force GUI update

        except Exception as e:
            print(f"Error updating hardware info: {str(e)}")

    def update_computer_list(self, computers):
        """Update the computer list panel"""
        self.computer_list.delete('1.0', tk.END)
        for computer in computers:
            self.computer_list.insert(tk.END, f"📱 {computer['hostname']}\n")
            self.computer_list.insert(tk.END, f"   OS: {computer['os']}\n")
            self.computer_list.insert(tk.END, f"   CPU Cores: {computer['cpu_count']}\n")
            self.computer_list.insert(tk.END, "\n")

    def update_network_info(self, data):
        """Update network monitoring displays"""
        # Update network statistics
        stats = data['io_counters']
        self.network_stats.delete('1.0', tk.END)
        self.network_stats.insert('1.0',
                                  f"Bytes Sent: {stats['bytes_sent']:,}\n"
                                  f"Bytes Received: {stats['bytes_recv']:,}\n"
                                  f"Packets Sent: {stats['packets_sent']:,}\n"
                                  f"Packets Received: {stats['packets_recv']:,}\n"
                                  f"Errors In: {stats['errin']}\n"
                                  f"Errors Out: {stats['errout']}\n"
                                  )

        # Update connections list
        self.connections_tree.delete(*self.connections_tree.get_children())
        for conn in data['connections']:
            if conn['laddr'] and conn['raddr']:  # Only show active connections
                self.connections_tree.insert('', 'end', values=(
                    f"{conn['laddr'][0]}:{conn['laddr'][1]}",
                    f"{conn['raddr'][0]}:{conn['raddr'][1]}",
                    conn['status']
                ))

    def refresh_software_list(self):
        """Refresh the installed software list"""
        response = self.send_command('software_inventory', {})
        if response['status'] == 'success':
            self.software_tree.delete(*self.software_tree.get_children())
            for software in response['data']:
                self.software_tree.insert('', 'end', values=(
                    software['name'],
                    software['version']
                ))

    def install_software(self):
        """Handle software installation"""
        file_path = filedialog.askopenfilename(
            filetypes=[("Executable files", "*.exe"), ("All files", "*.*")]
        )
        if file_path:
            response = self.send_command('software_install', {
                'file_path': file_path
            })
            if response['status'] == 'success':
                messagebox.showinfo("Installation", "Software installed successfully")
                self.refresh_software_list()
            else:
                messagebox.showerror("Installation Error", response['message'])

    def uninstall_software(self):
        """Handle software uninstallation"""
        selected = self.software_tree.selection()
        if not selected:
            messagebox.showwarning("Selection", "Please select software to uninstall")
            return

        software = self.software_tree.item(selected[0])['values'][0]
        if messagebox.askyesno("Confirm Uninstall", f"Uninstall {software}?"):
            response = self.send_command('software_uninstall', {
                'software': software
            })
            if response['status'] == 'success':
                messagebox.showinfo("Uninstallation", "Software uninstalled successfully")
                self.refresh_software_list()
            else:
                messagebox.showerror("Uninstallation Error", response['message'])

    def power_action(self, action):
        """Execute power management action"""
        if messagebox.askyesno("Confirm Action", f"Execute {action}?"):
            response = self.send_command('power_management', {
                'action': action
            })
            if response['status'] == 'success':
                messagebox.showinfo("Power Management", f"{action.title()} initiated")
            else:
                messagebox.showerror("Power Management Error", response['message'])

    def schedule_power_action(self):
        """Schedule a power management action"""
        time = self.schedule_time.get()
        action = self.power_action_var.get()

        try:
            # Validate time format
            hour, minute = map(int, time.split(':'))
            if not (0 <= hour <= 23 and 0 <= minute <= 59):
                raise ValueError()

            response = self.send_command('schedule_power', {
                'action': action,
                'time': time
            })

            if response['status'] == 'success':
                messagebox.showinfo("Schedule", f"{action.title()} scheduled for {time}")
            else:
                messagebox.showerror("Schedule Error", response['message'])

        except ValueError:
            messagebox.showerror("Invalid Time", "Please enter time in HH:MM format")

    def select_files(self):
        """Open file selection dialog"""
        files = filedialog.askopenfilenames()
        if files:
            self.source_path.delete(0, tk.END)
            self.source_path.insert(0, ';'.join(files))

    def select_destination(self):
        """Open destination selection dialog"""
        folder = filedialog.askdirectory()
        if folder:
            self.dest_path.delete(0, tk.END)
            self.dest_path.insert(0, folder)

    def transfer_files(self):
        """Handle file transfer"""
        source_files = self.source_path.get().split(';')
        destination = self.dest_path.get()

        if not source_files or not destination:
            messagebox.showwarning("Transfer", "Please select source and destination")
            return

        for file_path in source_files:
            try:
                response = self.send_command('file_transfer', {
                    'operation': 'send',
                    'source': file_path,
                    'destination': destination
                })

                if response['status'] == 'success':
                    self.transfer_log.insert('end', f"Transferred: {file_path}\n")
                else:
                    self.transfer_log.insert('end', f"Failed: {file_path} - {response['message']}\n")

                self.transfer_log.see('end')

            except Exception as e:
                self.transfer_log.insert('end', f"Error: {file_path} - {str(e)}\n")
                self.transfer_log.see('end')

    def execute_command(self):
        """Execute command on remote system"""
        command = self.command_input.get()
        if not command:
            return

        response = self.send_command('execute_command', {
            'command': command
        })

        self.command_output.delete('1.0', tk.END)
        if response['status'] == 'success':
            output = response['data']
            self.command_output.insert('1.0',
                                       f"Command: {command}\n"
                                       f"Output:\n{output['stdout']}\n"
                                       f"Errors:\n{output['stderr']}\n"
                                       f"Return Code: {output['return_code']}\n"
                                       )
        else:
            self.command_output.insert('1.0', f"Error: {response['message']}\n")

    def run_saved_command(self):
        """Execute saved command"""
        command = self.saved_commands.get()
        self.command_input.delete(0, tk.END)
        self.command_input.insert(0, command)
        self.execute_command()


if __name__ == "__main__":
    app = MCCClient()
    app.mainloop()
