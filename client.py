import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import socket
import json
import threading
from cryptography.fernet import Fernet
import customtkinter as ctk
import time
import os


class MCCClient(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Initialize connection management
        self.connections = {}  # Dictionary to store all connections
        self.active_connection = None  # Currently selected connection

        # Configure the window
        self.title("Multi Computers Control")
        self.geometry("1200x800")

        # Set the theme
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        self.create_gui()
        self.initialize_monitoring()

    def create_gui(self):
        """Create the main GUI with connection management"""
        # Create main container
        self.main_container = ctk.CTkFrame(self)
        self.main_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Create sidebar for computer list and connection management
        self.sidebar = ctk.CTkFrame(self.main_container, width=250)
        self.sidebar.pack(side=tk.LEFT, fill=tk.Y, padx=5, pady=5)

        # Add connection controls
        connection_frame = ctk.CTkFrame(self.sidebar)
        connection_frame.pack(fill=tk.X, padx=5, pady=5)

        self.host_entry = ctk.CTkEntry(connection_frame, placeholder_text="IP Address")
        self.host_entry.pack(fill=tk.X, padx=5, pady=2)

        self.port_entry = ctk.CTkEntry(connection_frame, placeholder_text="Port (default: 5000)")
        self.port_entry.pack(fill=tk.X, padx=5, pady=2)

        btn_frame = ctk.CTkFrame(connection_frame)
        btn_frame.pack(fill=tk.X, padx=5, pady=2)

        ctk.CTkButton(btn_frame, text="Connect", command=self.add_connection).pack(side=tk.LEFT, padx=2)
        ctk.CTkButton(btn_frame, text="Disconnect", command=self.remove_connection).pack(side=tk.LEFT, padx=2)

        # Create computer list with selection capability
        self.computer_list = ttk.Treeview(self.sidebar, columns=("status",), show="tree headings")
        self.computer_list.heading("#0", text="Computers")
        self.computer_list.heading("status", text="Status")
        self.computer_list.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.computer_list.bind('<<TreeviewSelect>>', self.on_computer_select)

        # Create main content area with tabs
        self.notebook = ttk.Notebook(self.main_container)
        self.notebook.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)

        self.create_monitoring_tab()
        self.create_software_tab()
        self.create_power_tab()
        self.create_file_transfer_tab()
        self.create_command_tab()
        self.create_network_tab()

    def create_monitoring_tab(self):
        """Create the monitoring tab with improved widget management"""
        monitoring_frame = ctk.CTkFrame(self.notebook)
        self.notebook.add(monitoring_frame, text="Monitoring")

        # Top section for CPU and Memory in a single frame
        top_section = ctk.CTkFrame(monitoring_frame)
        top_section.pack(fill=tk.X, padx=10, pady=5)

        # CPU Usage
        self.cpu_frame = ctk.CTkFrame(top_section)
        self.cpu_frame.pack(fill=tk.X, pady=5)
        self.cpu_label_text = ctk.CTkLabel(self.cpu_frame, text="CPU Usage:")
        self.cpu_label_text.pack(side=tk.LEFT, padx=5)
        self.cpu_progress = ctk.CTkProgressBar(self.cpu_frame)
        self.cpu_progress.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.cpu_progress.set(0)
        self.cpu_label = ctk.CTkLabel(self.cpu_frame, text="0%")
        self.cpu_label.pack(side=tk.LEFT, padx=5)

        # Memory Usage
        self.mem_frame = ctk.CTkFrame(top_section)
        self.mem_frame.pack(fill=tk.X, pady=5)
        self.mem_label_text = ctk.CTkLabel(self.mem_frame, text="Memory Usage:")
        self.mem_label_text.pack(side=tk.LEFT, padx=5)
        self.mem_progress = ctk.CTkProgressBar(self.mem_frame)
        self.mem_progress.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.mem_progress.set(0)
        self.mem_label = ctk.CTkLabel(self.mem_frame, text="0%")
        self.mem_label.pack(side=tk.LEFT, padx=5)

        # Container frame for disk usage
        self.disk_container = ctk.CTkFrame(monitoring_frame)
        self.disk_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Label for disk section
        self.disk_label = ctk.CTkLabel(self.disk_container, text="Disk Usage:")
        self.disk_label.pack(anchor=tk.W, padx=5, pady=5)

        # Scrollable frame for disk information
        self.disk_frame = ctk.CTkFrame(self.disk_container)
        self.disk_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

    def create_software_tab(self):
        """Create the software management tab with stable widget handling"""
        # Main container frame
        self.software_tab = ctk.CTkFrame(self.notebook)
        self.notebook.add(self.software_tab, text="Software")

        # Top status frame
        self.status_frame = ctk.CTkFrame(self.software_tab)
        self.status_frame.pack(fill=tk.X, padx=10, pady=(5, 0))
        self.status_label = ctk.CTkLabel(self.status_frame, text="Select a computer to view installed software")
        self.status_label.pack(fill=tk.X, padx=5, pady=5)

        # Software list frame
        self.software_list_frame = ctk.CTkFrame(self.software_tab)
        self.software_list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Create and pack the Treeview with scrollbar
        tree_container = ttk.Frame(self.software_list_frame)
        tree_container.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        self.software_tree = ttk.Treeview(
            tree_container,
            columns=("Name", "Version"),
            show="headings"
        )

        # Configure columns
        self.software_tree.heading("Name", text="Software Name")
        self.software_tree.heading("Version", text="Version")
        self.software_tree.column("Name", width=300, minwidth=200)
        self.software_tree.column("Version", width=100, minwidth=100)

        # Create and configure scrollbar
        scrollbar = ttk.Scrollbar(tree_container, orient="vertical", command=self.software_tree.yview)
        self.software_tree.configure(yscrollcommand=scrollbar.set)

        # Pack the tree and scrollbar
        self.software_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Button frame at bottom
        button_frame = ctk.CTkFrame(self.software_tab)
        button_frame.pack(fill=tk.X, padx=10, pady=(0, 5))

        # Create buttons using grid for better stability
        self.refresh_btn = ctk.CTkButton(
            button_frame,
            text="Refresh List",
            command=self.refresh_software_list,
            width=120
        )
        self.refresh_btn.pack(side=tk.LEFT, padx=5, pady=5)

        self.install_btn = ctk.CTkButton(
            button_frame,
            text="Install Software",
            command=self.install_software,
            width=120
        )
        self.install_btn.pack(side=tk.LEFT, padx=5, pady=5)

        self.uninstall_btn = ctk.CTkButton(
            button_frame,
            text="Uninstall",
            command=self.uninstall_software,
            width=120
        )
        self.uninstall_btn.pack(side=tk.LEFT, padx=5, pady=5)

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

    def create_network_tab(self):
        """Create the network monitoring tab"""
        network_frame = ctk.CTkFrame(self.notebook)
        self.notebook.add(network_frame, text="Network")

        # Network statistics
        self.network_stats = ctk.CTkTextbox(network_frame, height=100)
        self.network_stats.pack(fill=tk.X, padx=5, pady=5)

        # Active connections
        self.connections_tree = ttk.Treeview(network_frame,
                                             columns=("Local", "Remote", "Status"), show="headings")
        self.connections_tree.heading("Local", text="Local Address")
        self.connections_tree.heading("Remote", text="Remote Address")
        self.connections_tree.heading("Status", text="Status")
        self.connections_tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

    def add_connection(self):
        """Add a new remote connection"""
        host = self.host_entry.get()
        port = self.port_entry.get()

        if not host:
            messagebox.showwarning("Input Error", "Please enter an IP address")
            return

        try:
            port = int(port) if port else 5000
        except ValueError:
            messagebox.showwarning("Input Error", "Invalid port number")
            return

        try:
            # Create new connection
            new_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            new_socket.settimeout(5)  # 5 second timeout for connection
            new_socket.connect((host, port))

            # Receive encryption key
            encryption_key = new_socket.recv(1024)
            cipher_suite = Fernet(encryption_key)

            # Store connection info
            connection_id = f"{host}:{port}"
            self.connections[connection_id] = {
                'socket': new_socket,
                'cipher_suite': cipher_suite,
                'host': host,
                'port': port,
                'system_info': None
            }

            # Start monitoring thread for this connection
            thread = threading.Thread(
                target=self.monitor_connection,
                args=(connection_id,),
                daemon=True
            )
            thread.start()

            # Add to computer list
            self.computer_list.insert('', 'end', connection_id, text=host, values=('Connected',))

            # Clear input fields
            self.host_entry.delete(0, tk.END)
            self.port_entry.delete(0, tk.END)

        except Exception as e:
            messagebox.showerror("Connection Error", f"Failed to connect: {str(e)}")

    def remove_connection(self):
        """Remove selected connection"""
        selected = self.computer_list.selection()
        if not selected:
            messagebox.showwarning("Selection", "Please select a computer to disconnect")
            return

        connection_id = selected[0]
        if connection_id in self.connections:
            try:
                self.connections[connection_id]['socket'].close()
            except:
                pass
            del self.connections[connection_id]
            self.computer_list.delete(connection_id)

            if self.active_connection == connection_id:
                self.active_connection = None

    def on_computer_select(self, event):
        """Handle computer selection"""
        selected = self.computer_list.selection()
        if selected:
            self.active_connection = selected[0]
            self.refresh_monitoring()
        else:
            self.active_connection = None

    def monitor_connection(self, connection_id):
        """Monitor individual connection"""
        connection = self.connections.get(connection_id)
        if not connection:
            return

        while connection_id in self.connections:
            try:
                # Get system info
                response = self.send_command(connection_id, 'system_info', {})
                if response and response.get('status') == 'success':
                    self.connections[connection_id]['system_info'] = response['data']

                # Update status in computer list
                self.computer_list.set(connection_id, "status", "Connected")

            except Exception as e:
                print(f"Monitoring error for {connection_id}: {str(e)}")
                self.computer_list.set(connection_id, "status", "Error")
                break

            time.sleep(5)  # Check every 5 seconds

    def send_command(self, connection_id, command_type, data):
        """Send command to specific connection"""
        connection = self.connections.get(connection_id)
        if not connection:
            return None

        try:
            command = {
                'type': command_type,
                'data': data
            }

            encrypted_data = connection['cipher_suite'].encrypt(json.dumps(command).encode())
            connection['socket'].send(encrypted_data)

            encrypted_response = connection['socket'].recv(4096)
            response = json.loads(connection['cipher_suite'].decrypt(encrypted_response).decode())
            return response

        except Exception as e:
            print(f"Command error for {connection_id}: {str(e)}")
            return None

    def update_hardware_info(self, data):
        """Update hardware monitoring displays with Task Manager style disk info"""
        try:
            # Update CPU usage
            cpu_percent = data.get('cpu_percent', 0)
            self.cpu_progress.configure(mode="determinate")
            self.cpu_progress.set(cpu_percent / 100.0)
            self.cpu_label.configure(text=f"{cpu_percent:.1f}%")

            # Update memory usage
            memory_data = data.get('memory_usage', {})
            memory_percent = memory_data.get('percent', 0)
            self.mem_progress.configure(mode="determinate")
            self.mem_progress.set(memory_percent / 100.0)
            self.mem_label.configure(text=f"{memory_percent:.1f}%")

            # Clear existing disk information
            for widget in self.disk_frame.winfo_children():
                widget.destroy()

            # Create header
            header_frame = ctk.CTkFrame(self.disk_frame)
            header_frame.pack(fill=tk.X, padx=5, pady=(0, 5))

            headers = ["Drive", "Capacity", "Used Space", "Free Space", "Usage"]
            widths = [100, 150, 150, 150, 100]  # Approximate widths for each column

            for header, width in zip(headers, widths):
                label = ctk.CTkLabel(header_frame, text=header, width=width)
                label.pack(side=tk.LEFT, padx=5)

            # Update disk usage display with Task Manager style
            disk_usage = data.get('disk_usage', {})
            for mount, usage in disk_usage.items():
                try:
                    # Create frame for this disk
                    disk_frame = ctk.CTkFrame(self.disk_frame)
                    disk_frame.pack(fill=tk.X, padx=5, pady=2)

                    # Drive letter/name
                    ctk.CTkLabel(disk_frame, text=mount, width=100).pack(side=tk.LEFT, padx=5)

                    # Total capacity
                    total_gb = usage.get('total', 0) / (1024 ** 3)
                    ctk.CTkLabel(disk_frame,
                                 text=f"{total_gb:.1f} GB",
                                 width=150
                                 ).pack(side=tk.LEFT, padx=5)

                    # Used space
                    used_gb = usage.get('used', 0) / (1024 ** 3)
                    ctk.CTkLabel(disk_frame,
                                 text=f"{used_gb:.1f} GB",
                                 width=150
                                 ).pack(side=tk.LEFT, padx=5)

                    # Free space
                    free_gb = (usage.get('total', 0) - usage.get('used', 0)) / (1024 ** 3)
                    ctk.CTkLabel(disk_frame,
                                 text=f"{free_gb:.1f} GB",
                                 width=150
                                 ).pack(side=tk.LEFT, padx=5)

                    # Usage percentage and progress bar
                    percent_frame = ctk.CTkFrame(disk_frame)
                    percent_frame.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

                    percent = usage.get('percent', 0)
                    progress = ctk.CTkProgressBar(percent_frame)
                    progress.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
                    progress.set(percent / 100.0)

                    # Color the progress bar based on usage
                    if percent >= 90:
                        progress.configure(progress_color="red")
                    elif percent >= 75:
                        progress.configure(progress_color="orange")
                    else:
                        progress.configure(progress_color="green")

                    ctk.CTkLabel(percent_frame,
                                 text=f"{percent:.1f}%",
                                 width=50
                                 ).pack(side=tk.LEFT, padx=5)

                except Exception as disk_error:
                    print(f"Error displaying disk {mount}: {str(disk_error)}")
                    continue

        except Exception as e:
            print(f"Error updating hardware info: {str(e)}")

    def update_software_status(self, message):
        """Update status message safely"""
        try:
            if hasattr(self, 'status_label') and self.status_label.winfo_exists():
                self.status_label.configure(text=message)
                self.software_tab.update_idletasks()
        except Exception as e:
            print(f"Error updating status: {e}")

    def refresh_monitoring(self):
        """Refresh monitoring data with improved error handling"""
        if not self.active_connection:
            return

        try:
            response = self.send_command(self.active_connection, 'hardware_monitor', {})
            if response and isinstance(response, dict):
                if response.get('status') == 'success' and isinstance(response.get('data'), dict):
                    self.update_hardware_info(response['data'])
                else:
                    print("Invalid response format from server")
            else:
                print("No valid response from server")
        except Exception as e:
            print(f"Refresh error: {str(e)}")

    def install_software(self):
        """Safely handle software installation"""
        if not self.active_connection:
            self.update_software_status("Please select a computer first")
            return

        try:
            file_path = filedialog.askopenfilename(
                title="Select Software to Install",
                filetypes=[
                    ("Executable files", "*.exe"),
                    ("MSI files", "*.msi"),
                    ("All files", "*.*")
                ]
            )

            if file_path:
                self.update_software_status("Installing software... Please wait")

                with open(file_path, 'rb') as file:
                    file_data = file.read()
                    response = self.send_command(self.active_connection, 'software_install', {
                        'filename': os.path.basename(file_path),
                        'data': file_data.hex()
                    })

                if response and response.get('status') == 'success':
                    self.update_software_status("Software installed successfully")
                    self.refresh_software_list()
                else:
                    self.update_software_status("Installation failed")

        except Exception as e:
            self.update_software_status(f"Installation error: {str(e)}")

    def uninstall_software(self):
        """Safely handle software uninstallation"""
        if not self.active_connection:
            self.update_software_status("Please select a computer first")
            return

        selected = self.software_tree.selection()
        if not selected:
            self.update_software_status("Please select software to uninstall")
            return

        try:
            software = self.software_tree.item(selected[0])['values'][0]
            if messagebox.askyesno("Confirm Uninstall", f"Are you sure you want to uninstall {software}?"):
                self.update_software_status("Uninstalling software...")

                response = self.send_command(self.active_connection, 'software_uninstall', {
                    'software': software
                })

                if response and response.get('status') == 'success':
                    self.update_software_status("Software uninstalled successfully")
                    self.refresh_software_list()
                else:
                    self.update_software_status("Uninstallation failed")

        except Exception as e:
            self.update_software_status(f"Uninstallation error: {str(e)}")

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
        if not self.active_connection:
            messagebox.showwarning("Connection", "Please select a computer first")
            return

        source_files = self.source_path.get().split(';')
        destination = self.dest_path.get()

        if not source_files or not destination:
            messagebox.showwarning("Transfer", "Please select source and destination")
            return

        for file_path in source_files:
            try:
                with open(file_path, 'rb') as file:
                    file_data = file.read()
                    response = self.send_command(self.active_connection, 'file_transfer', {
                        'operation': 'send',
                        'filename': os.path.basename(file_path),
                        'destination': destination,
                        'data': file_data.decode('utf-8', errors='ignore')
                    })

                if response and response.get('status') == 'success':
                    self.transfer_log.insert('end', f"Transferred: {file_path}\n")
                else:
                    self.transfer_log.insert('end', f"Failed: {file_path}\n")

                self.transfer_log.see('end')

            except Exception as e:
                self.transfer_log.insert('end', f"Error: {file_path} - {str(e)}\n")
                self.transfer_log.see('end')

    def execute_command(self):
        """Execute command on remote system"""
        if not self.active_connection:
            messagebox.showwarning("Connection", "Please select a computer first")
            return

        command = self.command_input.get()
        if not command:
            return

        response = self.send_command(self.active_connection, 'execute_command', {
            'command': command
        })

        self.command_output.delete('1.0', tk.END)
        if response and response.get('status') == 'success':
            output = response['data']
            self.command_output.insert('1.0',
                                       f"Command: {command}\n"
                                       f"Output:\n{output['stdout']}\n"
                                       f"Errors:\n{output['stderr']}\n"
                                       f"Return Code: {output['return_code']}\n"
                                       )
        else:
            self.command_output.insert('1.0', f"Error executing command\n")

    def power_action(self, action):
        """Execute power management action"""
        if not self.active_connection:
            messagebox.showwarning("Connection", "Please select a computer first")
            return

        if messagebox.askyesno("Confirm Action", f"Execute {action} on the selected computer?"):
            response = self.send_command(self.active_connection, 'power_management', {
                'action': action
            })

            if response and response.get('status') == 'success':
                messagebox.showinfo("Power Management", f"{action.title()} initiated")
            else:
                messagebox.showerror("Power Management Error", "Failed to execute power action")

    def refresh_software_list(self):
        """Safely refresh the software list"""
        if not self.active_connection:
            self.update_software_status("Please select a computer first")
            return

        try:
            self.update_software_status("Retrieving software list...")

            response = self.send_command(self.active_connection, 'software_inventory', {})

            if response and response.get('status') == 'success':
                # Clear existing items
                for item in self.software_tree.get_children():
                    self.software_tree.delete(item)

                # Add new items
                for software in response['data']:
                    self.software_tree.insert('', 'end', values=(
                        software.get('name', 'Unknown'),
                        software.get('version', 'Unknown')
                    ))

                self.update_software_status("Software list updated successfully")
            else:
                self.update_software_status("Failed to retrieve software list")

        except Exception as e:
            self.update_software_status(f"Error refreshing list: {str(e)}")

    def initialize_monitoring(self):
        """Initialize monitoring thread"""
        self.monitoring_thread = threading.Thread(target=self.monitor_resources, daemon=True)
        self.monitoring_thread.start()

    def monitor_resources(self):
        """Monitor system resources"""
        while True:
            if self.active_connection:
                self.refresh_monitoring()
            time.sleep(1)


if __name__ == "__main__":
    app = MCCClient()
    app.mainloop()
