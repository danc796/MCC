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

        # widget state tracking
        self.active_tab = None
        self.monitoring_active = False
        self.progress_bars = {}

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
        self.create_remote_desktop_tab()

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
        """Create the software management tab"""
        self.software_tab = ctk.CTkFrame(self.notebook)
        self.notebook.add(self.software_tab, text="Software")

        # Top frame for status and search
        top_frame = ctk.CTkFrame(self.software_tab)
        top_frame.pack(fill=tk.X, padx=10, pady=(5, 0))

        # Status label
        self.status_label = ctk.CTkLabel(
            top_frame,
            text="Select a computer to view installed software"
        )
        self.status_label.pack(side=tk.LEFT, padx=5, pady=5)

        # Search frame
        search_frame = ctk.CTkFrame(top_frame)
        search_frame.pack(side=tk.RIGHT, padx=5, pady=5)

        # Search entry
        self.search_entry = ctk.CTkEntry(
            search_frame,
            placeholder_text="Search software...",
            width=200
        )
        self.search_entry.pack(side=tk.LEFT, padx=(5, 0))

        # Bind search entry to search function
        self.search_entry.bind('<KeyRelease>', self.on_search)

        # Clear search button
        self.clear_search_btn = ctk.CTkButton(
            search_frame,
            text="Clear",
            command=self.clear_search,
            width=60
        )
        self.clear_search_btn.pack(side=tk.LEFT, padx=5)

        # Software list frame
        self.software_list_frame = ctk.CTkFrame(self.software_tab)
        self.software_list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Create and pack the Treeview with scrollbar
        tree_container = ttk.Frame(self.software_list_frame)
        tree_container.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Create Treeview
        self.software_tree = ttk.Treeview(
            tree_container,
            columns=("Name", "Version"),
            show="headings"
        )

        # Configure columns
        self.software_tree.heading("Name", text="Software Name",
                                   command=lambda: self.treeview_sort_column("Name", False))
        self.software_tree.heading("Version", text="Version",
                                   command=lambda: self.treeview_sort_column("Version", False))

        self.software_tree.column("Name", width=400, minwidth=200)
        self.software_tree.column("Version", width=150, minwidth=100)

        # Create and configure scrollbar
        scrollbar = ttk.Scrollbar(tree_container, orient="vertical", command=self.software_tree.yview)
        self.software_tree.configure(yscrollcommand=scrollbar.set)

        # Pack the tree and scrollbar
        self.software_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Refresh button
        self.refresh_btn = ctk.CTkButton(
            self.software_tab,
            text="Refresh List",
            command=self.refresh_software_list,
            width=120
        )
        self.refresh_btn.pack(pady=5)

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

    def create_remote_desktop_tab(self):  # Renamed from create_network_tab
        """Create the remote desktop tab"""
        remote_desktop_frame = ctk.CTkFrame(self.notebook)
        self.notebook.add(remote_desktop_frame, text="Remote Desktop")  # Changed text from "Network"

        # Network statistics
        stats_frame = ctk.CTkFrame(remote_desktop_frame)
        stats_frame.pack(fill=tk.X, padx=5, pady=5)

        # Add labels for network stats
        self.network_sent_label = ctk.CTkLabel(stats_frame, text="Sent: 0 B")
        self.network_sent_label.pack(side=tk.LEFT, padx=10)

        self.network_recv_label = ctk.CTkLabel(stats_frame, text="Received: 0 B")
        self.network_recv_label.pack(side=tk.LEFT, padx=10)

        # Active connections
        conn_frame = ctk.CTkFrame(remote_desktop_frame)
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
            remote_desktop_frame,
            text="Refresh",
            command=self.refresh_network_info
        ).pack(pady=5)

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
        """Send command with improved debugging"""
        connection = self.connections.get(connection_id)
        if not connection:
            print(f"No connection found for ID: {connection_id}")
            return None

        try:
            # Prepare command
            command = {
                'type': command_type,
                'data': data
            }
            print(f"Sending command: {command_type}")

            # Encrypt and send
            encrypted_data = connection['cipher_suite'].encrypt(json.dumps(command).encode())
            connection['socket'].send(encrypted_data)
            print("Command sent successfully")

            # Receive response
            print("Waiting for response...")
            encrypted_response = connection['socket'].recv(16384)
            if not encrypted_response:
                print("Received empty response from server")
                return None
            print(f"Received encrypted response of length: {len(encrypted_response)}")

            # Decrypt response
            try:
                decrypted_response = connection['cipher_suite'].decrypt(encrypted_response).decode()
                print("Response decrypted successfully")
            except Exception as e:
                print(f"Decryption error: {str(e)}")
                return None

            # Parse JSON
            try:
                response = json.loads(decrypted_response)
                print(f"Response parsed successfully: {response.get('status', 'unknown status')}")
                return response
            except json.JSONDecodeError as e:
                print(f"JSON parsing error: {str(e)}")
                return None

        except socket.timeout:
            print(f"Connection timeout for {connection_id}")
            return None
        except ConnectionError as e:
            print(f"Connection error for {connection_id}: {str(e)}")
            return None
        except Exception as e:
            print(f"Unexpected error: {str(e)}")
            return None

    def update_hardware_info(self, data):
        """Update hardware monitoring displays with widget validation"""
        if not self.monitoring_active:
            return

        try:
            # Update CPU usage with validation
            if 'cpu' in self.progress_bars and self.progress_bars['cpu'].winfo_exists():
                cpu_percent = data.get('cpu_percent', 0)
                if isinstance(cpu_percent, (int, float)) and 0 <= cpu_percent <= 100:
                    self.progress_bars['cpu'].set(cpu_percent / 100.0)
                    if hasattr(self, 'cpu_label') and self.cpu_label.winfo_exists():
                        self.cpu_label.configure(text=f"{cpu_percent:.1f}%")

            # Update memory usage with validation
            if 'mem' in self.progress_bars and self.progress_bars['mem'].winfo_exists():
                memory_data = data.get('memory_usage', {})
                if isinstance(memory_data, dict):
                    memory_percent = memory_data.get('percent', 0)
                    if isinstance(memory_percent, (int, float)) and 0 <= memory_percent <= 100:
                        self.progress_bars['mem'].set(memory_percent / 100.0)
                        if hasattr(self, 'mem_label') and self.mem_label.winfo_exists():
                            self.mem_label.configure(text=f"{memory_percent:.1f}%")

            # Only update disk info if we're on the monitoring tab
            if self.monitoring_active and hasattr(self, 'disk_frame') and self.disk_frame.winfo_exists():
                # Clear existing disk information
                for widget in self.disk_frame.winfo_children():
                    widget.destroy()

                # Update disk usage
                self.update_disk_info(data.get('disk_usage', {}))

        except Exception as e:
            print(f"Error updating hardware info: {str(e)}")

    def update_disk_info(self, disk_usage):
        """Separate method for updating disk information"""
        try:
            if not isinstance(disk_usage, dict) or not self.monitoring_active:
                return

            # Create header
            header_frame = ctk.CTkFrame(self.disk_frame)
            header_frame.pack(fill=tk.X, padx=5, pady=(0, 5))

            headers = ["Drive", "Capacity", "Used Space", "Free Space", "Usage"]
            widths = [100, 150, 150, 150, 100]

            for header, width in zip(headers, widths):
                ctk.CTkLabel(header_frame, text=header, width=width).pack(side=tk.LEFT, padx=5)

            for mount, usage in disk_usage.items():
                if not isinstance(usage, dict):
                    continue

                try:
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

                    # Usage percentage
                    percent_frame = ctk.CTkFrame(disk_frame)
                    percent_frame.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

                    if self.monitoring_active and self.disk_frame.winfo_exists():
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
                    print(f"Error displaying disk {mount}: {str(disk_error)}")
                    continue

        except Exception as e:
            print(f"Error updating disk info: {str(e)}")

    def update_software_status(self, message):
        """Update status message in software tab"""
        if hasattr(self, 'status_label') and self.status_label.winfo_exists():
            self.status_label.configure(text=message)

    def on_search(self, event=None):
        """Handle software search"""
        if not self.active_connection:
            self.update_software_status("Please select a computer first")
            return

        search_term = self.search_entry.get().strip()

        # Clear the tree
        for item in self.software_tree.get_children():
            self.software_tree.delete(item)

        try:
            response = self.send_command(
                self.active_connection,
                'software_inventory',
                {'search': ''}  # Get all software
            )

            if response and response.get('status') == 'success':
                software_list = response.get('data', [])

                # Filter software list based on search term
                filtered_list = []
                for software in software_list:
                    if isinstance(software, dict):
                        name = software.get('name', '').lower()
                        version = software.get('version', '').lower()
                        if search_term.lower() in name or search_term.lower() in version:
                            filtered_list.append(software)

                # Update the tree with filtered results
                for software in filtered_list:
                    self.software_tree.insert('', 'end', values=(
                        software.get('name', 'Unknown'),
                        software.get('version', 'N/A')
                    ))

                # Update status
                status = f"Found {len(filtered_list)} software items"
                if search_term:
                    status += f" matching '{search_term}'"
                self.update_software_status(status)

        except Exception as e:
            self.update_software_status(f"Error during search: {str(e)}")

    def clear_search(self):
        """Clear search and refresh list"""
        self.search_entry.delete(0, tk.END)
        self.refresh_software_list()  # Refresh with no search term

    def treeview_sort_column(self, col, reverse):
        """Sort treeview column"""
        l = [(self.software_tree.set(k, col), k) for k in self.software_tree.get_children('')]
        l.sort(reverse=reverse)

        # Rearrange items in sorted positions
        for index, (val, k) in enumerate(l):
            self.software_tree.move(k, '', index)

        # Reverse sort next time
        self.software_tree.heading(col, command=lambda: self.treeview_sort_column(col, not reverse))

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

    def refresh_software_list(self, search_term=""):
        """Refresh software list"""
        if not self.active_connection:
            self.update_software_status("Please select a computer first")
            return

        try:
            self.update_software_status("Retrieving software list...")

            # Clear existing items
            for item in self.software_tree.get_children():
                self.software_tree.delete(item)

            # Set timeout for software inventory
            connection = self.connections.get(self.active_connection)
            if not connection:
                self.update_software_status("Connection not found")
                return

            # Set temporary longer timeout
            original_timeout = connection['socket'].gettimeout()
            connection['socket'].settimeout(30)

            try:
                response = self.send_command(
                    self.active_connection,
                    'software_inventory',
                    {'search': search_term}
                )

                if not response:
                    self.update_software_status("No response from server")
                    return

                if response.get('status') == 'error':
                    self.update_software_status(f"Server error: {response.get('message', 'Unknown error')}")
                    return

                if response.get('status') == 'success':
                    software_list = response.get('data', [])

                    for software in software_list:
                        if isinstance(software, dict):
                            name = software.get('name', 'Unknown')
                            version = software.get('version', 'N/A')

                            self.software_tree.insert('', 'end', values=(name, version))

                    status = f"Found {len(software_list)} software items"
                    if search_term:
                        status += f" matching '{search_term}'"
                    self.update_software_status(status)
                else:
                    self.update_software_status("Invalid response format")

            finally:
                # Restore original timeout
                connection['socket'].settimeout(original_timeout)

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
        """Monitor system resources with tab awareness"""
        while True:
            try:
                if self.active_connection and self.monitoring_active:
                    if hasattr(self, 'cpu_progress') and hasattr(self, 'mem_progress'):
                        self.refresh_monitoring()
                time.sleep(2)  # Reduced update frequency
            except Exception as e:
                print(f"Monitor resources error: {str(e)}")
                time.sleep(2)

    def on_tab_change(self, event):
        """Handle tab changes with widget state management"""
        try:
            current_tab = self.notebook.select()
            prev_tab = self.active_tab
            self.active_tab = self.notebook.tab(current_tab, "text")

            print(f"Tab changed from {prev_tab} to {self.active_tab}")

            # Handle monitoring tab exit
            if prev_tab == "Monitoring":
                self.monitoring_active = False
                # Clear progress bars dict
                self.progress_bars.clear()

            # Handle monitoring tab entry
            if self.active_tab == "Monitoring":
                self.monitoring_active = True
                # Store progress bar references
                if hasattr(self, 'cpu_progress'):
                    self.progress_bars['cpu'] = self.cpu_progress
                if hasattr(self, 'mem_progress'):
                    self.progress_bars['mem'] = self.mem_progress

            # Update the UI based on the new tab
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
