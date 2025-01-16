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

        # Store reference to main frame
        self.monitoring_main_frame = monitoring_frame

        # Add cleanup binding when tab changes
        self.notebook.bind('<<NotebookTabChanged>>', self.on_tab_change)

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
        """Create the power management tab with theme-aware buttons"""
        power_frame = ctk.CTkFrame(self.notebook)
        self.notebook.add(power_frame, text="Power")

        # Title and status
        title_frame = ctk.CTkFrame(power_frame)
        title_frame.pack(pady=20)

        ctk.CTkLabel(
            title_frame,
            text="Power Management",
            font=("Helvetica", 20)
        ).pack()

        # Status indicator
        self.power_status = ctk.CTkLabel(
            title_frame,
            text="Select a computer to manage power options",
            font=("Helvetica", 12)
        )
        self.power_status.pack(pady=10)

        # Button container
        button_frame = ctk.CTkFrame(power_frame)
        button_frame.pack(expand=True)

        # Create shutdown button
        shutdown_btn = ctk.CTkButton(
            button_frame,
            text="Shutdown",
            command=lambda: self.power_action_with_confirmation(
                "shutdown",
                "This will shut down the remote computer. Continue?"
            ),
            width=200,
            height=40,
            hover_color="#FF6B6B"  # Red hover
        )
        shutdown_btn.pack(pady=10)

        # Create restart button
        restart_btn = ctk.CTkButton(
            button_frame,
            text="Restart",
            command=lambda: self.power_action_with_confirmation(
                "restart",
                "This will restart the remote computer. Continue?"
            ),
            width=200,
            height=40,
            hover_color="#4D96FF"  # Blue hover
        )
        restart_btn.pack(pady=10)

        # Create sleep button
        sleep_btn = ctk.CTkButton(
            button_frame,
            text="Sleep",
            command=lambda: self.power_action_with_confirmation(
                "sleep",
                "This will put the remote computer to sleep. Continue?"
            ),
            width=200,
            height=40,
            hover_color="#6BCB77"  # Green hover
        )
        sleep_btn.pack(pady=10)

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
        stats_frame = ctk.CTkFrame(network_frame)
        stats_frame.pack(fill=tk.X, padx=5, pady=5)

        # Add labels for network stats
        self.network_sent_label = ctk.CTkLabel(stats_frame, text="Sent: 0 B")
        self.network_sent_label.pack(side=tk.LEFT, padx=10)

        self.network_recv_label = ctk.CTkLabel(stats_frame, text="Received: 0 B")
        self.network_recv_label.pack(side=tk.LEFT, padx=10)

        # Active connections
        conn_frame = ctk.CTkFrame(network_frame)
        conn_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Add header label
        ctk.CTkLabel(conn_frame, text="Active Network Connections:").pack(anchor=tk.W, padx=5, pady=5)

        # Create Treeview with scrollbar
        tree_frame = ttk.Frame(conn_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        self.connections_tree = ttk.Treeview(
            tree_frame,
            columns=("Local", "Remote", "Status", "Type"),
            show="headings"
        )

        # Configure columns
        self.connections_tree.heading("Local", text="Local Address")
        self.connections_tree.heading("Remote", text="Remote Address")
        self.connections_tree.heading("Status", text="Status")
        self.connections_tree.heading("Type", text="Type")

        # Set column widths
        self.connections_tree.column("Local", width=200)
        self.connections_tree.column("Remote", width=200)
        self.connections_tree.column("Status", width=100)
        self.connections_tree.column("Type", width=100)

        # Add scrollbar
        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.connections_tree.yview)
        self.connections_tree.configure(yscrollcommand=scrollbar.set)

        # Pack tree and scrollbar
        self.connections_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Add refresh button
        ctk.CTkButton(
            network_frame,
            text="Refresh",
            command=self.refresh_network_info
        ).pack(pady=5)

        def create_power_button(self, parent, text, action, confirm_msg, hover_color):
            """Create a power management button with hover effect and confirmation"""
            button = ctk.CTkButton(
                parent,
                text=text,
                command=lambda: self.power_action_with_confirmation(action, confirm_msg),
                width=200,
                height=40,
                fg_color=("#333333", "#2B2B2B"),  # Dark theme colors
                hover_color=hover_color
            )
            return button

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
        """Send command with improved connection handling"""
        connection = self.connections.get(connection_id)
        if not connection:
            return None

        try:
            # Set a reasonable timeout
            connection['socket'].settimeout(5)

            command = {
                'type': command_type,
                'data': data
            }

            encrypted_data = connection['cipher_suite'].encrypt(json.dumps(command).encode())
            connection['socket'].send(encrypted_data)

            encrypted_response = connection['socket'].recv(16384)  # Increased buffer size
            if not encrypted_response:
                raise ConnectionError("Empty response from server")

            decrypted_response = connection['cipher_suite'].decrypt(encrypted_response).decode()
            response = json.loads(decrypted_response)

            return response

        except socket.timeout:
            print(f"Connection timeout for {connection_id}")
            return None
        except ConnectionError as e:
            print(f"Connection error for {connection_id}: {str(e)}")
            return None
        except json.JSONDecodeError as e:
            print(f"Invalid JSON response from {connection_id}: {str(e)}")
            return None
        except Exception as e:
            print(f"Command error for {connection_id}: {str(e)}")
            return None
        finally:
            # Reset timeout
            try:
                connection['socket'].settimeout(None)
            except:
                pass

    def update_hardware_info(self, data):
        """Update hardware monitoring displays with widget existence checks"""
        try:
            # Check if widgets still exist before updating
            if not hasattr(self, 'cpu_progress') or not self.cpu_progress.winfo_exists():
                return
            if not hasattr(self, 'mem_progress') or not self.mem_progress.winfo_exists():
                return
            if not hasattr(self, 'disk_frame') or not self.disk_frame.winfo_exists():
                return

            # Update CPU usage with validation
            cpu_percent = data.get('cpu_percent', 0)
            if isinstance(cpu_percent, (int, float)) and 0 <= cpu_percent <= 100:
                self.cpu_progress.set(cpu_percent / 100.0)
                if hasattr(self, 'cpu_label') and self.cpu_label.winfo_exists():
                    self.cpu_label.configure(text=f"{cpu_percent:.1f}%")

            # Update memory usage with validation
            memory_data = data.get('memory_usage', {})
            if isinstance(memory_data, dict):
                memory_percent = memory_data.get('percent', 0)
                if isinstance(memory_percent, (int, float)) and 0 <= memory_percent <= 100:
                    self.mem_progress.set(memory_percent / 100.0)
                    if hasattr(self, 'mem_label') and self.mem_label.winfo_exists():
                        self.mem_label.configure(text=f"{memory_percent:.1f}%")

            # Safely clear existing disk information
            try:
                for widget in self.disk_frame.winfo_children():
                    if widget.winfo_exists():
                        widget.destroy()
            except Exception:
                return

            # Update disk usage with validation
            disk_usage = data.get('disk_usage', {})
            if isinstance(disk_usage, dict) and self.disk_frame.winfo_exists():
                try:
                    # Create header
                    header_frame = ctk.CTkFrame(self.disk_frame)
                    header_frame.pack(fill=tk.X, padx=5, pady=(0, 5))

                    headers = ["Drive", "Capacity", "Used Space", "Free Space", "Usage"]
                    widths = [100, 150, 150, 150, 100]

                    for header, width in zip(headers, widths):
                        label = ctk.CTkLabel(header_frame, text=header, width=width)
                        label.pack(side=tk.LEFT, padx=5)

                    for mount, usage in disk_usage.items():
                        if not isinstance(usage, dict) or not self.disk_frame.winfo_exists():
                            continue

                        disk_frame = ctk.CTkFrame(self.disk_frame)
                        disk_frame.pack(fill=tk.X, padx=5, pady=2)

                        total = usage.get('total', 0)
                        used = usage.get('used', 0)
                        percent = usage.get('percent', 0)

                        if not all(isinstance(x, (int, float)) for x in [total, used, percent]):
                            continue

                        # Drive letter/name
                        ctk.CTkLabel(disk_frame, text=str(mount), width=100).pack(side=tk.LEFT, padx=5)

                        # Total capacity
                        total_gb = total / (1024 ** 3)
                        ctk.CTkLabel(disk_frame, text=f"{total_gb:.1f} GB", width=150).pack(side=tk.LEFT, padx=5)

                        # Used space
                        used_gb = used / (1024 ** 3)
                        ctk.CTkLabel(disk_frame, text=f"{used_gb:.1f} GB", width=150).pack(side=tk.LEFT, padx=5)

                        # Free space
                        free_gb = (total - used) / (1024 ** 3)
                        ctk.CTkLabel(disk_frame, text=f"{free_gb:.1f} GB", width=150).pack(side=tk.LEFT, padx=5)

                        # Usage percentage and progress bar
                        if self.disk_frame.winfo_exists():
                            percent_frame = ctk.CTkFrame(disk_frame)
                            percent_frame.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

                            progress = ctk.CTkProgressBar(percent_frame)
                            progress.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
                            progress.set(percent / 100.0)

                            # Color based on usage
                            if percent >= 90:
                                progress.configure(progress_color="red")
                            elif percent >= 75:
                                progress.configure(progress_color="orange")
                            else:
                                progress.configure(progress_color="green")

                            ctk.CTkLabel(percent_frame, text=f"{percent:.1f}%", width=50).pack(side=tk.LEFT, padx=5)

                except Exception as disk_error:
                    print(f"Error displaying disk info: {str(disk_error)}")

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

            # Early return if no response
            if not response:
                print("No response from server")
                return

            # Validate response format
            if not isinstance(response, dict):
                print("Invalid response format from server")
                return

            if response.get('status') != 'success':
                print(f"Error from server: {response.get('message', 'Unknown error')}")
                return

            data = response.get('data')
            if not isinstance(data, dict):
                print("Invalid data format from server")
                return

            # Update hardware info with validation
            self.update_hardware_info(data)

        except Exception as e:
            print(f"Refresh monitoring error: {str(e)}")

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

    def power_action_with_confirmation(self, action, confirm_msg):
        """Execute power action with confirmation and proper error handling"""
        if not self.active_connection:
            if hasattr(self, 'power_status') and self.power_status.winfo_exists():
                self.power_status.configure(
                    text="Please select a computer first",
                    text_color="red"
                )
            return

        if not self.connections.get(self.active_connection):
            if hasattr(self, 'power_status') and self.power_status.winfo_exists():
                self.power_status.configure(
                    text="Connection lost. Please reconnect.",
                    text_color="red"
                )
            return

        # Show confirmation dialog
        if messagebox.askyesno("Confirm Action", confirm_msg):
            try:
                if hasattr(self, 'power_status') and self.power_status.winfo_exists():
                    self.power_status.configure(
                        text=f"Initiating {action}...",
                        text_color="white"
                    )

                response = self.send_command(self.active_connection, 'power_management', {
                    'action': action
                })

                if hasattr(self, 'power_status') and self.power_status.winfo_exists():
                    if response and response.get('status') == 'success':
                        self.power_status.configure(
                            text=f"{action.capitalize()} command sent successfully",
                            text_color="green"
                        )
                    else:
                        self.power_status.configure(
                            text=f"Failed to execute {action}",
                            text_color="red"
                        )

            except Exception as e:
                if hasattr(self, 'power_status') and self.power_status.winfo_exists():
                    self.power_status.configure(
                        text=f"Error: {str(e)}",
                        text_color="red"
                    )
                print(f"Power action error: {str(e)}")

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

    def refresh_network_info(self):
        """Refresh network information with improved error handling"""
        if not self.active_connection:
            return

        try:
            response = self.send_command(self.active_connection, 'network_monitor', {})

            # Check if we got an error response
            if not response:
                return
            if response.get('status') == 'error':
                print(f"Error getting network info: {response.get('message', 'Unknown error')}")
                return

            if response.get('status') == 'success':
                data = response.get('data', {})

                # Update IO counters
                io_counters = data.get('io_counters', {})
                if io_counters:
                    bytes_sent = self.format_bytes(io_counters.get('bytes_sent', 0))
                    bytes_recv = self.format_bytes(io_counters.get('bytes_recv', 0))

                    self.network_sent_label.configure(text=f"Sent: {bytes_sent}")
                    self.network_recv_label.configure(text=f"Received: {bytes_recv}")

                # Clear existing items
                for item in self.connections_tree.get_children():
                    self.connections_tree.delete(item)

                # Update connections
                for conn in data.get('connections', []):
                    try:
                        laddr = conn.get('laddr', ('0.0.0.0', 0))
                        raddr = conn.get('raddr', ('0.0.0.0', 0))

                        # Format addresses
                        local = f"{laddr[0]}:{laddr[1]}" if isinstance(laddr, tuple) else "N/A"
                        remote = f"{raddr[0]}:{raddr[1]}" if isinstance(raddr, tuple) else "N/A"

                        status = conn.get('status', 'UNKNOWN')
                        type_ = 'TCP' if conn.get('type', socket.SOCK_STREAM) == socket.SOCK_STREAM else 'UDP'

                        self.connections_tree.insert('', 'end', values=(local, remote, status, type_))
                    except Exception as e:
                        print(f"Error processing connection: {str(e)}")
                        continue

        except Exception as e:
            print(f"Error refreshing network info: {str(e)}")

    def format_bytes(self, bytes_value):
        """Format bytes into human readable format with error handling"""
        try:
            bytes_num = float(bytes_value)
            for unit in ['B', 'KB', 'MB', 'GB']:
                if bytes_num < 1024.0:
                    return f"{bytes_num:.1f} {unit}"
                bytes_num /= 1024.0
            return f"{bytes_num:.1f} TB"
        except (TypeError, ValueError):
            return "0 B"

    def initialize_monitoring(self):
        """Initialize monitoring thread"""
        self.monitoring_thread = threading.Thread(target=self.monitor_resources, daemon=True)
        self.monitoring_thread.start()

    def monitor_resources(self):
        """Monitor system resources with improved widget checks"""
        while True:
            try:
                # Check if the main window still exists
                if not self.winfo_exists():
                    break

                if self.active_connection and self.active_connection in self.connections:
                    # Only refresh if the monitoring tab widgets exist
                    if (hasattr(self, 'cpu_progress') and
                            hasattr(self, 'mem_progress') and
                            hasattr(self, 'disk_frame') and
                            all(widget.winfo_exists() for widget in
                                [self.cpu_progress, self.mem_progress, self.disk_frame])):
                        self.refresh_monitoring()
                    time.sleep(2)
                else:
                    time.sleep(1)
            except Exception as e:
                print(f"Monitor resources error: {str(e)}")
                time.sleep(2)

    def on_tab_change(self, event):
        """Handle tab changes with improved widget management"""
        try:
            current_tab = self.notebook.select()
            tab_name = self.notebook.tab(current_tab, "text")

            # Reset status labels when leaving tabs
            if hasattr(self, 'power_status') and self.power_status.winfo_exists():
                self.power_status.configure(
                    text="Select a computer to manage power options",
                    text_color="white"
                )

            # Clean up monitoring tab widgets
            if tab_name != "Monitoring" and hasattr(self, 'disk_frame'):
                if self.disk_frame.winfo_exists():
                    for widget in self.disk_frame.winfo_children():
                        if widget.winfo_exists():
                            widget.destroy()

            # Force update the new tab
            self.update_idletasks()

        except Exception as e:
            print(f"Error during tab change: {str(e)}")

    def safe_widget_destroy(self, widget):
        """Safely destroy a widget if it exists"""
        try:
            if hasattr(self, widget) and getattr(self, widget).winfo_exists():
                getattr(self, widget).destroy()
        except Exception:
            pass

    def safe_widget_update(self, widget, **kwargs):
        """Safely update a widget's properties if it exists"""
        try:
            if hasattr(self, widget) and getattr(self, widget).winfo_exists():
                getattr(self, widget).configure(**kwargs)
                return True
        except Exception:
            pass
        return False

    def check_widget_exists(self, widget):
        """Safely check if a widget exists and is valid"""
        try:
            return widget.winfo_exists()
        except Exception:
            return False


if __name__ == "__main__":
    app = MCCClient()
    app.mainloop()
