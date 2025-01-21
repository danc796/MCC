import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import customtkinter as ctk
import socket
import json
import threading
from cryptography.fernet import Fernet
import time
import logging
from datetime import datetime


class ToastNotification:
    def __init__(self, parent):
        self.parent = parent
        self.notifications = []
        self.showing = False

    def show_toast(self, message, category="info"):
        """Show a toast notification with auto-dismiss"""
        # Configure colors based on category
        colors = {
            "success": "#28a745",  # Green
            "error": "#dc3545",  # Red
            "info": "#007bff",  # Blue
            "warning": "#ffc107"  # Yellow
        }
        bg_color = colors.get(category, "#6c757d")  # Default gray

        # Create a notification window
        toast = ctk.CTkFrame(self.parent)
        toast.configure(fg_color=bg_color)

        # Add a message
        label = ctk.CTkLabel(
            toast,
            text=message,
            text_color="white",
            font=("Helvetica", 12)
        )
        label.pack(padx=20, pady=10)

        # Position in bottom right
        screen_width = self.parent.winfo_width()
        screen_height = self.parent.winfo_height()

        # Add to notifications queue
        self.notifications.append({
            'widget': toast,
            'start_time': self.parent.after(0, lambda: None)  # Current time
        })

        # Show notification if not already showing
        if not self.showing:
            self._show_next_notification()

    def _show_next_notification(self):
        """Show the next notification in queue"""
        if not self.notifications:
            self.showing = False
            return

        self.showing = True
        notification = self.notifications[0]
        toast = notification['widget']

        # Position toast
        toast.place(
            relx=1,
            rely=1,
            anchor="se",
            x=-20,
            y=-20
        )

        # Schedule removal
        self.parent.after(3000, lambda: self._remove_notification(notification))

    def _remove_notification(self, notification):
        """Remove a notification and show next if any"""
        if notification in self.notifications:
            self.notifications.remove(notification)
            notification['widget'].destroy()

            # Show the next notification if any
            self._show_next_notification()


class MCCClient(ctk.CTk):
    def __init__(self):
        """Initialize the client with proper flag initialization"""
        # Initialize as CTk
        super().__init__()

        # Set up flags and state FIRST
        self.running = True  # Flag for monitoring threads
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

        # Set up logging
        logging.basicConfig(
            filename='mcc_client.log',
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s'
        )

        # Create GUI elements
        self.create_gui()

        # Set up shutdown handler
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

        # Initialize a toast notification system
        self.toast = ToastNotification(self)

        # Initialize monitoring LAST
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
        """Create enhanced power management tab"""
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

        self.power_status = ctk.CTkLabel(
            title_frame,
            text="Select computer(s) to manage power options",
            font=("Helvetica", 12)
        )
        self.power_status.pack(pady=10)

        # Action modes
        mode_frame = ctk.CTkFrame(power_frame)
        mode_frame.pack(pady=10)

        self.power_mode = tk.StringVar(value="single")

        single_radio = ctk.CTkRadioButton(
            mode_frame,
            text="Single Computer",
            variable=self.power_mode,
            value="single"
        )
        single_radio.pack(side=tk.LEFT, padx=10)

        all_radio = ctk.CTkRadioButton(
            mode_frame,
            text="All Computers",
            variable=self.power_mode,
            value="all"
        )
        all_radio.pack(side=tk.LEFT, padx=10)

        # Scheduled shutdown frame
        schedule_frame = ctk.CTkFrame(power_frame)
        schedule_frame.pack(pady=10, padx=20, fill=tk.X)

        ctk.CTkLabel(
            schedule_frame,
            text="Schedule Shutdown",
            font=("Helvetica", 14)
        ).pack(pady=5)

        # Center-aligned time input frame
        time_frame = ctk.CTkFrame(schedule_frame)
        time_frame.pack(pady=5)  # Removed fill=tk.X to allow centering

        # Time entry (HH:MM format) with label
        time_label = ctk.CTkLabel(
            time_frame,
            text="Enter time (HH:MM):",
            font=("Helvetica", 12)
        )
        time_label.pack(side=tk.LEFT, padx=5)

        self.schedule_time = ctk.CTkEntry(
            time_frame,
            placeholder_text="HH:MM",
            width=100
        )
        self.schedule_time.pack(side=tk.LEFT, padx=5)

        schedule_btn = ctk.CTkButton(
            schedule_frame,
            text="Schedule Shutdown",
            command=self.schedule_shutdown,
            width=200,
            height=40
        )
        schedule_btn.pack(pady=5)

        cancel_schedule_btn = ctk.CTkButton(
            schedule_frame,
            text="Cancel Scheduled Shutdown",
            command=lambda: self.power_action_with_confirmation(
                "cancel_scheduled",
                "Cancel all scheduled shutdowns?"
            ),
            width=200,
            height=40
        )
        cancel_schedule_btn.pack(pady=5)

        # Immediate actions frame
        actions_frame = ctk.CTkFrame(power_frame)
        actions_frame.pack(pady=10)

        # Create immediate action buttons
        buttons_data = [
            ("Shutdown", "shutdown", "This will shut down the selected computer(s). Continue?", "#FF6B6B"),
            ("Restart", "restart", "This will restart the selected computer(s). Continue?", "#4D96FF"),
            ("Lock Screen", "lock", "This will lock the selected computer(s). Continue?", "#FFB562")
        ]

        for text, action, confirm_msg, hover_color in buttons_data:
            ctk.CTkButton(
                actions_frame,
                text=text,
                command=lambda a=action, m=confirm_msg: self.power_action_with_confirmation(a, m),
                width=200,
                height=40,
                hover_color=hover_color
            ).pack(pady=5)

    def create_remote_desktop_tab(self):
        """Create the remote desktop tab with centered controls"""
        remote_desktop_frame = ctk.CTkFrame(self.notebook)
        self.notebook.add(remote_desktop_frame, text="Remote Desktop")

        # Title and message section
        title_frame = ctk.CTkFrame(remote_desktop_frame)
        title_frame.pack(pady=20)

        # Title
        ctk.CTkLabel(
            title_frame,
            text="Remote Desktop Control",
            font=("Helvetica", 20)
        ).pack()

        # Status message
        self.rdt_status = ctk.CTkLabel(
            title_frame,
            text="Select computer and activate the remote desktop",
            font=("Helvetica", 12)
        )
        self.rdt_status.pack(pady=10)

        # Center frame for the button
        center_frame = ctk.CTkFrame(remote_desktop_frame)
        center_frame.pack(expand=True)

        # Activation button
        self.rdt_button = ctk.CTkButton(
            center_frame,
            text="Activate Remote Desktop",
            # command=self.rdt_run,
            width=200,
            height=40
        )
        self.rdt_button.pack(pady=20)

        # Network statistics
        stats_frame = ctk.CTkFrame(remote_desktop_frame)
        stats_frame.pack(fill=tk.X, padx=5, pady=5)

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
        """Send command with improved large data handling"""
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

            # Convert to JSON and check size
            json_data = json.dumps(command)

            # Regular command sending
            encrypted_data = connection['cipher_suite'].encrypt(json_data.encode())
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
                return json.loads(decrypted_response)
            except Exception as e:
                print(f"Decryption error: {str(e)}")
                return None

        except Exception as e:
            print(f"Send command error: {str(e)}")
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

    def update_power_status(self, message, color="white"):
        """Update power status label with proper error handling"""
        try:
            if hasattr(self, 'power_status') and self.power_status.winfo_exists():
                self.power_status.configure(text=message, text_color=color)
        except Exception as e:
            logging.error(f"Error updating power status: {str(e)}")

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

    def power_action_with_confirmation(self, action, confirm_msg):
        """Execute power action with confirmation for single or multiple computers"""
        try:
            if self.power_mode.get() == "all":
                if not self.connections:
                    self.update_power_status("No computers connected", "red")
                    return

                if messagebox.askyesno("Confirm Action", f"{confirm_msg} (All Computers)"):
                    failed_computers = []
                    for conn_id in list(self.connections.keys()):  # Create a copy of keys
                        try:
                            response = self.send_command(conn_id, 'power_management', {
                                'action': action
                            })
                            if not response or response.get('status') != 'success':
                                failed_computers.append(self.connections[conn_id]['host'])
                        except Exception as e:
                            failed_computers.append(self.connections[conn_id]['host'])

                    if failed_computers:
                        self.update_power_status(f"Action failed for: {', '.join(failed_computers)}", "red")
                    else:
                        self.update_power_status(f"{action.capitalize()} initiated for all computers", "green")
            else:
                # Single computer mode
                if not self.active_connection:
                    self.update_power_status("Please select a computer first", "red")
                    return

                if messagebox.askyesno("Confirm Action", confirm_msg):
                    response = self.send_command(self.active_connection, 'power_management', {
                        'action': action
                    })

                    if response and response.get('status') == 'success':
                        self.update_power_status(f"{action.capitalize()} command sent successfully", "green")
                    else:
                        self.update_power_status(f"Failed to execute {action}", "red")

        except Exception as e:
            self.update_power_status(f"Error: {str(e)}", "red")
            logging.error(f"Power action error: {str(e)}")

    def schedule_shutdown(self):
        """Schedule a shutdown for the selected computer(s)"""
        try:
            time_str = self.schedule_time.get()

            if not time_str:
                self.update_power_status("Please enter time in HH:MM format", "red")
                return

            try:
                # Parse the time
                hours, minutes = map(int, time_str.split(':'))
                if not (0 <= hours <= 23 and 0 <= minutes <= 59):
                    raise ValueError("Invalid time values")

                # Calculate seconds until shutdown
                current_time = datetime.now()
                target_time = current_time.replace(hour=hours, minute=minutes, second=0, microsecond=0)

                # If the time has already passed today, schedule for tomorrow
                if target_time <= current_time:
                    target_time = target_time.replace(day=current_time.day + 1)

                seconds_until_shutdown = int((target_time - current_time).total_seconds())

                if self.power_mode.get() == "all":
                    if messagebox.askyesno("Confirm Action", "Schedule shutdown for all computers?"):
                        failed_computers = []
                        for conn_id in list(self.connections.keys()):
                            try:
                                response = self.send_command(conn_id, 'power_management', {
                                    'action': 'shutdown',
                                    'seconds': seconds_until_shutdown
                                })
                                if not response or response.get('status') != 'success':
                                    failed_computers.append(self.connections[conn_id]['host'])
                            except Exception:
                                failed_computers.append(self.connections[conn_id]['host'])

                        if failed_computers:
                            self.update_power_status(f"Scheduling failed for: {', '.join(failed_computers)}", "red")
                        else:
                            self.update_power_status("Shutdown scheduled for all computers", "green")
                else:
                    if not self.active_connection:
                        self.update_power_status("Please select a computer first", "red")
                        return

                    response = self.send_command(self.active_connection, 'power_management', {
                        'action': 'shutdown',
                        'seconds': seconds_until_shutdown
                    })

                    if response and response.get('status') == 'success':
                        self.update_power_status("Shutdown scheduled successfully", "green")
                    else:
                        self.update_power_status("Failed to schedule shutdown", "red")

            except ValueError:
                self.update_power_status("Invalid time format. Use HH:MM", "red")
        except Exception as e:
            self.update_power_status(f"Error: {str(e)}", "red")
            logging.error(f"Schedule shutdown error: {str(e)}")

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

    def initialize_monitoring(self):
        """Initialize monitoring thread safely"""
        if hasattr(self, 'monitoring_thread'):
            return  # Don't create multiple threads

        self.monitoring_thread = threading.Thread(
            target=self.monitor_resources,
            daemon=True  # Make it a daemon thread
        )
        self.monitoring_thread.start()
        logging.info("Monitoring thread started")

    def monitor_resources(self):
        """Monitor system resources with shutdown check"""
        logging.info("Starting resource monitoring")
        while getattr(self, 'running', True):  # Safe attribute access
            try:
                if not hasattr(self, 'active_connection'):
                    time.sleep(2)
                    continue

                if self.active_connection and self.monitoring_active:
                    if hasattr(self, 'cpu_progress') and hasattr(self, 'mem_progress'):
                        self.refresh_monitoring()
                time.sleep(2)  # Reduced update frequency
            except Exception as e:
                logging.error(f"Monitor resources error: {str(e)}")
                time.sleep(2)  # Wait before trying again

        logging.info("Resource monitoring stopped")

    def on_closing(self):
        """Handle window closing event"""
        try:
            # Set flag to stop monitoring threads
            self.running = False

            # Close all connections
            for conn_id in list(self.connections.keys()):
                try:
                    self.connections[conn_id]['socket'].close()
                except:
                    pass

            logging.info("Closing all connections")
            self.connections.clear()

            # Wait for threads to finish
            if hasattr(self, 'monitoring_thread') and self.monitoring_thread.is_alive():
                self.monitoring_thread.join(timeout=1.0)

            logging.info("Application shutting down")
            self.quit()

        except Exception as e:
            logging.error(f"Error during shutdown: {str(e)}")
            self.quit()

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


def format_bytes(bytes_value):
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


if __name__ == "__main__":
    app = MCCClient()
    app.mainloop()
