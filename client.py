import tkinter as tk
from tkinter import ttk, messagebox
import customtkinter as ctk
import socket
import json
import threading
from cryptography.fernet import Fernet
import time
import logging
from datetime import datetime
import sys
import cv2
import numpy as np
from PIL import Image, ImageTk
import struct


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
        super().__init__()

        self.running = True
        self.active_tab = None
        self.monitoring_active = False
        self.progress_bars = {}

        self.connections = {}
        self.active_connection = None

        self.last_cpu_percent = None
        self.last_memory_percent = None
        self.last_disk_usage = {}

        self.title("Multi Computers Control")
        self.geometry("1200x800")

        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        logging.basicConfig(
            filename='mcc_client.log',
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s'
        )

        # Initialize RDP attributes
        self.rdp_active = False
        self.rdp_socket = None
        self.rdp_display_thread = None
        self.rdp_connection = None

        self.create_gui()
        self.software_loaded_for = None  # Track which connection has loaded software

        self.protocol("WM_DELETE_WINDOW", self.on_closing)

        self.toast = ToastNotification(self)

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
        """Create a simplified remote desktop tab"""
        remote_desktop_frame = ctk.CTkFrame(self.notebook)
        self.notebook.add(remote_desktop_frame, text="Remote Desktop")

        # Top controls frame
        top_frame = ctk.CTkFrame(remote_desktop_frame)
        top_frame.pack(fill=tk.X, padx=10, pady=5)

        # Title and status
        title_frame = ctk.CTkFrame(top_frame)
        title_frame.pack(side=tk.LEFT, fill=tk.Y)

        ctk.CTkLabel(
            title_frame,
            text="Remote Desktop Control",
            font=("Helvetica", 16, "bold")
        ).pack(side=tk.TOP, anchor="w", padx=5)

        self.rdt_status = ctk.CTkLabel(
            title_frame,
            text="Select computer and activate the remote desktop",
            font=("Helvetica", 12)
        )
        self.rdt_status.pack(side=tk.TOP, anchor="w", padx=5, pady=2)

        # Create main display frame that takes most of the space, leaving room for buttons at bottom
        main_display_container = ctk.CTkFrame(remote_desktop_frame)
        main_display_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Create the RDP display frame inside the main container
        self.rdp_display_frame = ctk.CTkFrame(main_display_container, fg_color="#242424")
        self.rdp_display_frame.pack(fill=tk.BOTH, expand=True)

        # Create a canvas for the RDP display
        self.rdp_canvas = tk.Canvas(
            self.rdp_display_frame,
            bg="black",
            highlightthickness=0,
            borderwidth=0
        )
        self.rdp_canvas.pack(fill=tk.BOTH, expand=True)

        # Create a message for when no RDP is active
        self.rdp_message = ctk.CTkLabel(
            self.rdp_canvas,
            text="Remote Desktop Viewer\nClick 'Start Remote Desktop' to connect",
            font=("Helvetica", 16),
            text_color="white"
        )
        self.rdp_message.place(relx=0.5, rely=0.5, anchor=tk.CENTER)

        # Create a persistent control panel that will always be visible
        self.rdp_control_panel = ctk.CTkFrame(remote_desktop_frame, height=50)
        self.rdp_control_panel.pack(fill=tk.X, side=tk.BOTTOM, padx=10, pady=5)
        # Make sure this frame maintains its height
        self.rdp_control_panel.pack_propagate(False)

        # Control buttons - now in a persistent panel
        self.rdp_button = ctk.CTkButton(
            self.rdp_control_panel,
            text="Start Remote Desktop",
            command=self.toggle_rdp,
            width=150,
            height=30
        )
        self.rdp_button.pack(side=tk.LEFT, padx=5, pady=10)

        self.rdp_close_button = ctk.CTkButton(
            self.rdp_control_panel,
            text="Close RDP",
            command=self.stop_rdp,
            width=100,
            height=30,
            fg_color="#FF5555",
            hover_color="#FF0000",
            state="disabled"
        )
        self.rdp_close_button.pack(side=tk.LEFT, padx=5, pady=10)

        # Status indicator on the right
        self.rdp_status_indicator = ctk.CTkLabel(
            self.rdp_control_panel,
            text="Not connected",
            font=("Helvetica", 12),
            text_color="#888888"
        )
        self.rdp_status_indicator.pack(side=tk.RIGHT, padx=10, pady=10)

        # Store state for RDP
        self.rdp_active = False
        self.rdp_socket = None
        self.rdp_display_thread = None
        self.rdp_connection = None

        # Track if this tab is currently active to control keyboard input
        self.rdp_tab_active = False

        # Add bindings to track when this tab becomes active/inactive
        self.notebook.bind('<<NotebookTabChanged>>', self.on_tab_change)

    def add_connection(self):
        """Add a new remote connection with auto-reconnect support"""
        host = self.host_entry.get()
        port = self.port_entry.get()

        if not host:
            self.toast.show_toast("Please enter an IP address", "warning")
            return

        try:
            port = int(port) if port else 5000
            if port < 0 or port > 65535:
                raise ValueError("Port out of range")
        except ValueError:
            self.toast.show_toast("Invalid port number. Please enter a number between 0-65535", "warning")
            return

        connection_id = f"{host}:{port}"

        # Check if this connection already exists
        if connection_id in self.connections:
            self.toast.show_toast(f"Connection to {host}:{port} already exists", "warning")
            return

        # Store connection info first (even before successful connection)
        self.connections[connection_id] = {
            'socket': None,
            'cipher_suite': None,
            'host': host,
            'port': port,
            'system_info': None,
            'connection_active': False,
            'reconnect_thread': None
        }

        # Add to computer list with 'Connecting' status
        self.computer_list.insert('', 'end', connection_id, text=host, values=('Connecting...',))

        # Start connection in a separate thread to avoid UI freezing
        connect_thread = threading.Thread(
            target=self.connect_to_server,
            args=(connection_id,)
        )
        connect_thread.daemon = True
        connect_thread.start()

        # Clear input fields
        self.host_entry.delete(0, tk.END)
        self.port_entry.delete(0, tk.END)

    def connect_to_server(self, connection_id):
        """Connect to server with timeout handling"""
        connection = self.connections.get(connection_id)
        if not connection:
            return

        host, port = connection['host'], connection['port']

        try:
            # Create socket with timeout
            client_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            client_socket.settimeout(3.0)  # 3-second timeout

            # Attempt connection
            client_socket.connect((host, port))

            # Get encryption key
            key = client_socket.recv(44)  # Fernet key length
            cipher_suite = Fernet(key)

            # Update connection info
            self.connections[connection_id]['socket'] = client_socket
            self.connections[connection_id]['cipher_suite'] = cipher_suite
            self.connections[connection_id]['connection_active'] = True

            # Update UI from the main thread
            self.after(0, lambda: self.computer_list.set(connection_id, "status", "Connected"))

            # Start monitoring thread
            monitor_thread = threading.Thread(
                target=self.monitor_connection,
                args=(connection_id,)
            )
            monitor_thread.daemon = True
            monitor_thread.start()

        except (socket.timeout, ConnectionRefusedError) as e:
            # Handle timeouts and connection refusals
            error_msg = "Connection timed out" if isinstance(e, socket.timeout) else "Connection refused"

            # Update UI from the main thread
            self.after(0, lambda: self.computer_list.set(connection_id, "status", "Failed"))
            self.after(0, lambda: self.toast.show_toast(f"{error_msg} for {host}:{port}", "error"))

        except Exception as e:
            # Handle other errors
            self.after(0, lambda: self.computer_list.set(connection_id, "status", "Error"))
            self.after(0, lambda: self.toast.show_toast(f"Connection error: {str(e)}", "error"))

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
            # Check if active connection is changing
            if self.active_connection != selected[0]:
                self.active_connection = selected[0]
                # Reset software loaded state when connection changes
                self.software_loaded_for = None
                # If software tab is active, load software for new connection
                if self.active_tab == "Software":
                    self.auto_load_software()
                # Refresh monitoring for new connection
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
        """Update hardware monitoring displays with widget validation and value preservation"""
        if not self.monitoring_active:
            return

        try:
            # Update CPU usage with validation
            if 'cpu' in self.progress_bars and self.progress_bars['cpu'].winfo_exists():
                cpu_percent = data.get('cpu_percent')

                # Only update if value is not None and greater than 0
                if cpu_percent is not None and isinstance(cpu_percent, (int, float)) and cpu_percent > 0:
                    self.progress_bars['cpu'].set(cpu_percent / 100.0)
                    if hasattr(self, 'cpu_label') and self.cpu_label.winfo_exists():
                        self.cpu_label.configure(text=f"{cpu_percent:.1f}%")

            # Update memory usage with validation
            if 'mem' in self.progress_bars and self.progress_bars['mem'].winfo_exists():
                memory_data = data.get('memory_usage', {})
                if isinstance(memory_data, dict):
                    memory_percent = memory_data.get('percent')

                    # Only update if value is not None and greater than 0
                    if memory_percent is not None and isinstance(memory_percent, (int, float)) and memory_percent > 0:
                        self.progress_bars['mem'].set(memory_percent / 100.0)
                        if hasattr(self, 'mem_label') and self.mem_label.winfo_exists():
                            self.mem_label.configure(text=f"{memory_percent:.1f}%")

            # Only update disk info if we're on the monitoring tab and there's valid data
            if (self.monitoring_active and hasattr(self, 'disk_frame') and
                    self.disk_frame.winfo_exists() and 'disk_usage' in data and data['disk_usage']):

                # Don't clear existing disk information if there's no valid data
                disk_usage = data.get('disk_usage', {})
                if disk_usage and isinstance(disk_usage, dict) and any(disk_usage.values()):
                    # Only then update disk info
                    self.update_disk_info(disk_usage)

        except Exception as e:
            print(f"Error updating hardware info: {str(e)}")

    def update_disk_info(self, disk_usage):
        """Separate method for updating disk information with reduced flicker"""
        try:
            if not isinstance(disk_usage, dict) or not self.monitoring_active:
                return

            # Check if we already have disks displayed and new data is similar
            existing_disks = [w for w in self.disk_frame.winfo_children() if isinstance(w, ctk.CTkFrame)]

            # Only recreate UI if data structure has changed significantly
            if len(existing_disks) > 1:  # Header + at least one disk row
                drives = disk_usage.keys()

                # If we have roughly the same number of drives, update in place
                if len(drives) == len(existing_disks) - 1:  # -1 for header
                    self._update_disk_values(disk_usage, existing_disks)
                    return

            # Clear existing disk information
            for widget in self.disk_frame.winfo_children():
                widget.destroy()

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

    def _update_disk_values(self, disk_usage, existing_frames):
        """Update disk values in-place without destroying and recreating widgets"""
        try:
            # Skip header frame (first frame)
            header_frame = existing_frames[0]
            disk_frames = existing_frames[1:]

            # Update each disk frame with new values if match is found
            for i, (mount, usage) in enumerate(disk_usage.items()):
                if not isinstance(usage, dict) or i >= len(disk_frames):
                    continue

                disk_frame = disk_frames[i]
                labels = [w for w in disk_frame.winfo_children() if isinstance(w, ctk.CTkLabel)]

                # If structure doesn't match, skip this update
                if len(labels) < 2:
                    continue

                try:
                    total = usage.get('total', 0)
                    used = usage.get('used', 0)
                    percent = usage.get('percent', 0)

                    if not all(isinstance(x, (int, float)) for x in [total, used, percent]):
                        continue

                    # Drive name should stay the same
                    # [0] = Drive name label

                    # Update values only
                    # [1] = Total capacity
                    total_gb = total / (1024 ** 3)
                    if len(labels) > 1:
                        labels[1].configure(text=f"{total_gb:.1f} GB")

                    # [2] = Used space
                    used_gb = used / (1024 ** 3)
                    if len(labels) > 2:
                        labels[2].configure(text=f"{used_gb:.1f} GB")

                    # [3] = Free space
                    free_gb = (total - used) / (1024 ** 3)
                    if len(labels) > 3:
                        labels[3].configure(text=f"{free_gb:.1f} GB")

                    # Update percentage frames
                    percent_frames = [w for w in disk_frame.winfo_children()
                                      if isinstance(w, ctk.CTkFrame) and w != header_frame]

                    if percent_frames:
                        percent_frame = percent_frames[0]

                        # Get progress bar
                        progress_bars = [w for w in percent_frame.winfo_children()
                                         if isinstance(w, ctk.CTkProgressBar)]
                        if progress_bars:
                            progress = progress_bars[0]
                            progress.set(percent / 100.0)

                            # Update color based on usage
                            if percent >= 90:
                                progress.configure(progress_color="red")
                            elif percent >= 75:
                                progress.configure(progress_color="orange")
                            else:
                                progress.configure(progress_color="green")

                        # Update percentage label
                        percent_labels = [w for w in percent_frame.winfo_children()
                                          if isinstance(w, ctk.CTkLabel)]
                        if percent_labels:
                            percent_labels[0].configure(text=f"{percent:.1f}%")

                except Exception as disk_error:
                    print(f"Error updating disk values for {mount}: {str(disk_error)}")

        except Exception as e:
            print(f"Error in _update_disk_values: {str(e)}")

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
        """Handle software search with case-insensitive matching"""
        if not self.active_connection:
            self.update_software_status("Please select a computer first")
            return

        search_term = self.search_entry.get().strip().lower()  # Convert search term to lowercase

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

                # Filter software list based on search term - case insensitive
                filtered_list = []
                for software in software_list:
                    if isinstance(software, dict):
                        name = software.get('name', '').lower()  # Convert name to lowercase
                        version = software.get('version', '').lower()  # Convert version to lowercase
                        if search_term in name or search_term in version:
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
                    status += f" matching '{self.search_entry.get().strip()}'"  # Show original search term
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
        """Refresh monitoring data with improved error handling and value preservation"""
        if not self.active_connection:
            return

        try:
            response = self.send_command(self.active_connection, 'hardware_monitor', {})

            # Early return if no response - don't reset values
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

    def toggle_rdp(self):
        """Toggle RDP session on/off"""
        if not self.rdp_active:
            self.start_rdp()
        else:
            self.stop_rdp()

    def start_rdp(self):
        """Start integrated RDP session with enhanced UI feedback"""
        if not self.active_connection:
            self.rdt_status.configure(text="Please select a computer first")
            return

        try:
            # Request RDP server start
            response = self.send_command(self.active_connection, 'start_rdp', {})

            if response and response.get('status') == 'success':
                ip, port = response['data']['ip'], response['data']['port']

                # Update status with clear feedback
                self.rdt_status.configure(text="Connecting to remote desktop...")

                # Update status in the control panel
                if hasattr(self, 'rdp_status_indicator') and self.rdp_status_indicator.winfo_exists():
                    self.rdp_status_indicator.configure(
                        text=f"Connecting to {self.connections[self.active_connection]['host']}...",
                        text_color="#FFAA00"
                    )

                self.rdp_button.configure(text="Connecting...", state="disabled")

                # Create a socket
                self.rdp_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
                self.rdp_socket.settimeout(5.0)
                self.rdp_socket.connect((ip, port))

                # Send platform info
                platform_code = b'win' if sys.platform == "win32" else b'osx' if sys.platform == "darwin" else b'x11'
                self.rdp_socket.sendall(platform_code)

                # Hide the welcome message
                if hasattr(self, 'rdp_message') and self.rdp_message.winfo_exists():
                    self.rdp_message.place_forget()

                # Set up display thread
                self.rdp_active = True
                self.rdp_connection = self.active_connection
                self.rdp_display_thread = threading.Thread(target=self.rdp_display_loop)
                self.rdp_display_thread.daemon = True
                self.rdp_display_thread.start()

                # Set up input handlers
                self.setup_rdp_input_handlers()

                # Update UI with successful connection status
                self.rdt_status.configure(text="Remote desktop session active")

                # Update status in the control panel
                if hasattr(self, 'rdp_status_indicator') and self.rdp_status_indicator.winfo_exists():
                    self.rdp_status_indicator.configure(
                        text=f"Connected to {self.connections[self.active_connection]['host']}",
                        text_color="#00CC00"  # Green for active connection
                    )

                self.rdp_button.configure(text="Connected", state="disabled")
                self.rdp_close_button.configure(state="normal")
            else:
                self.rdt_status.configure(text="Failed to start remote desktop")

                # Update status in the control panel
                if hasattr(self, 'rdp_status_indicator') and self.rdp_status_indicator.winfo_exists():
                    self.rdp_status_indicator.configure(
                        text="Connection failed",
                        text_color="#FF0000"  # Red for error
                    )

                self.rdp_button.configure(text="Start Remote Desktop", state="normal")
        except Exception as e:
            self.rdt_status.configure(text=f"RDP Error: {str(e)}")

            # Update status in the control panel
            if hasattr(self, 'rdp_status_indicator') and self.rdp_status_indicator.winfo_exists():
                self.rdp_status_indicator.configure(
                    text="Connection error",
                    text_color="#FF0000"  # Red for error
                )

            logging.error(f"RDP error: {str(e)}")
            self.rdp_button.configure(text="Start Remote Desktop", state="normal")
            self.stop_rdp()

    def stop_rdp(self):
        """Simplified stop RDP function"""
        try:
            self.rdp_active = False

            # Close socket
            if self.rdp_socket:
                try:
                    self.rdp_socket.close()
                except:
                    pass
                self.rdp_socket = None

            # Clear display
            if hasattr(self, 'rdp_canvas') and self.rdp_canvas.winfo_exists():
                self.rdp_canvas.delete("all")
                # Show welcome message again
                if hasattr(self, 'rdp_message') and self.rdp_message.winfo_exists():
                    self.rdp_message.place(relx=0.5, rely=0.5, anchor=tk.CENTER)

            # Stop RDP server
            if self.rdp_connection:
                self.send_command(self.rdp_connection, 'stop_rdp', {})
                self.rdp_connection = None

            # Update status
            if hasattr(self, 'rdt_status') and self.rdt_status.winfo_exists():
                self.rdt_status.configure(text="Remote desktop session ended")

            # Update status indicator
            if hasattr(self, 'rdp_status_indicator') and self.rdp_status_indicator.winfo_exists():
                self.rdp_status_indicator.configure(text="Not connected", text_color="#888888")

            # Update buttons - always reset to initial state
            if hasattr(self, 'rdp_button') and self.rdp_button.winfo_exists():
                self.rdp_button.configure(text="Start Remote Desktop", state="normal")

            if hasattr(self, 'rdp_close_button') and self.rdp_close_button.winfo_exists():
                self.rdp_close_button.configure(state="disabled")

        except Exception as e:
            logging.error(f"Error stopping RDP: {str(e)}")

            # Force reset of critical UI elements even if there was an error
            if hasattr(self, 'rdp_button') and self.rdp_button.winfo_exists():
                self.rdp_button.configure(text="Start Remote Desktop", state="normal")

    def receive_rdp_frame(self):
        """Receive a frame from the RDP server"""
        # Get header
        header = self.receive_exact(5)
        img_type, length = struct.unpack(">BI", header)

        # Get image data
        img_data = b''
        buffer_size = 10240
        while length > 0:
            chunk_size = min(buffer_size, length)
            chunk = self.receive_exact(chunk_size)
            img_data += chunk
            length -= len(chunk)

        return img_type, img_data

    def receive_exact(self, size):
        """Receive exact number of bytes"""
        data = b''
        while len(data) < size and self.rdp_active:
            try:
                chunk = self.rdp_socket.recv(size - len(data))
                if not chunk:
                    raise ConnectionError("Connection lost")
                data += chunk
            except socket.timeout:
                continue
            except Exception as e:
                logging.error(f"Error receiving data: {str(e)}")
                raise
        return data

    def rdp_display_loop(self):
        """Display loop for RDP with proper scaling and positioning"""
        try:
            last_image = None

            while self.rdp_active:
                try:
                    img_type, img_data = self.receive_rdp_frame()

                    # Process image
                    np_arr = np.frombuffer(img_data, dtype=np.uint8)
                    img = cv2.imdecode(np_arr, cv2.IMREAD_COLOR)

                    if img_type == 0 and last_image is not None:  # Diff frame
                        img = cv2.bitwise_xor(last_image, img)

                    last_image = img.copy()

                    # Convert to PIL format
                    img_rgb = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
                    pil_img = Image.fromarray(img_rgb)

                    # Update canvas if it exists
                    if hasattr(self, 'rdp_canvas') and self.rdp_canvas.winfo_exists():
                        # Store original image dimensions
                        original_width, original_height = pil_img.size

                        # Get canvas dimensions
                        canvas_width = self.rdp_canvas.winfo_width()
                        canvas_height = self.rdp_canvas.winfo_height()

                        # Only resize if canvas dimensions are valid and different
                        # from the image dimensions
                        if (canvas_width > 10 and canvas_height > 10 and
                                (original_width != canvas_width or original_height != canvas_height)):
                            # Calculate scale factors for both dimensions
                            width_scale = canvas_width / original_width
                            height_scale = canvas_height / original_height

                            # Use the smaller scale factor to ensure the image fits
                            scale = min(width_scale, height_scale)

                            # Calculate new dimensions
                            new_width = int(original_width * scale)
                            new_height = int(original_height * scale)

                            # Resize image
                            pil_img = pil_img.resize((new_width, new_height), Image.Resampling.LANCZOS)

                        # Convert to PhotoImage
                        photo_img = ImageTk.PhotoImage(image=pil_img)

                        # Calculate position to center the image in the canvas
                        x_position = (canvas_width - pil_img.width) // 2
                        y_position = (canvas_height - pil_img.height) // 2

                        # Clear existing content
                        self.rdp_canvas.delete("all")

                        # Create new image
                        self.rdp_canvas.create_image(x_position, y_position, anchor=tk.NW, image=photo_img)

                        # Keep reference to prevent garbage collection
                        self.rdp_canvas.photo = photo_img

                except socket.timeout:
                    # Timeout is expected, just continue
                    continue
                except ConnectionError:
                    break
                except Exception as e:
                    logging.error(f"Error in RDP display loop: {str(e)}")
                    break

        except Exception as e:
            logging.error(f"RDP display error: {str(e)}")
        finally:
            # Clean up on exit
            self.after(0, self.stop_rdp)

    def setup_rdp_input_handlers(self):
        """Set up input handlers for RDP with tab-aware keyboard handling"""
        if not hasattr(self, 'rdp_canvas') or not self.rdp_canvas.winfo_exists():
            return

        # Mouse constants
        MOUSE_LEFT = 201
        MOUSE_SCROLL = 202
        MOUSE_RIGHT = 203
        MOUSE_MOVE = 204

        # Mouse event handling is direct and always active when over the canvas
        self.rdp_canvas.bind("<Button-1>", lambda e: self.send_rdp_mouse_event(MOUSE_LEFT, 100, e.x, e.y))
        self.rdp_canvas.bind("<ButtonRelease-1>", lambda e: self.send_rdp_mouse_event(MOUSE_LEFT, 117, e.x, e.y))
        self.rdp_canvas.bind("<Button-3>", lambda e: self.send_rdp_mouse_event(MOUSE_RIGHT, 100, e.x, e.y))
        self.rdp_canvas.bind("<ButtonRelease-3>", lambda e: self.send_rdp_mouse_event(MOUSE_RIGHT, 117, e.x, e.y))

        # Add a scale factor for mouse position to handle resolution differences
        def scale_mouse_position(event):
            if not hasattr(self, 'rdp_canvas') or not self.rdp_canvas.winfo_exists():
                return

            # Get the actual image dimensions if we have them
            if hasattr(self.rdp_canvas, 'photo') and self.rdp_canvas.photo:
                img_width = self.rdp_canvas.photo.width()
                img_height = self.rdp_canvas.photo.height()
            else:
                # Default to canvas dimensions if no image
                img_width = self.rdp_canvas.winfo_width()
                img_height = self.rdp_canvas.winfo_height()

            # Get canvas dimensions
            canvas_width = self.rdp_canvas.winfo_width()
            canvas_height = self.rdp_canvas.winfo_height()

            if canvas_width <= 1 or canvas_height <= 1 or img_width <= 1 or img_height <= 1:
                # Avoid division by zero
                self.send_rdp_mouse_event(MOUSE_MOVE, 0, event.x, event.y)
                return

            # Calculate the scale factors
            x_scale = img_width / canvas_width
            y_scale = img_height / canvas_height

            # Scale the coordinates
            scaled_x = int(event.x * x_scale)
            scaled_y = int(event.y * y_scale)

            # Send the scaled coordinates
            self.send_rdp_mouse_event(MOUSE_MOVE, 0, scaled_x, scaled_y)

        # Use the scaled position function for mouse movement
        self.rdp_canvas.bind("<Motion>", scale_mouse_position)

        # Ensure canvas gets focus when clicked
        self.rdp_canvas.bind("<Button>", lambda e: self.rdp_canvas.focus_set())

        # Mouse wheel
        if sys.platform in ("win32", "darwin"):
            self.rdp_canvas.bind("<MouseWheel>",
                                 lambda e: self.send_rdp_mouse_event(MOUSE_SCROLL,
                                                                     1 if e.delta > 0 else 0,
                                                                     e.x, e.y))
        else:
            self.rdp_canvas.bind("<Button-4>",
                                 lambda e: self.send_rdp_mouse_event(MOUSE_SCROLL, 1, e.x, e.y))
            self.rdp_canvas.bind("<Button-5>",
                                 lambda e: self.send_rdp_mouse_event(MOUSE_SCROLL, 0, e.x, e.y))

        # Keyboard events are only processed when RDP tab is active
        def handle_key_press(event):
            if self.rdp_tab_active and self.rdp_active:
                self.send_rdp_key_event(event.keysym, 100)

        def handle_key_release(event):
            if self.rdp_tab_active and self.rdp_active:
                self.send_rdp_key_event(event.keysym, 117)

        # Bind keyboard events to the canvas and also to the main window
        self.rdp_canvas.bind("<KeyPress>", handle_key_press)
        self.rdp_canvas.bind("<KeyRelease>", handle_key_release)

        # Set focus to the canvas
        self.rdp_canvas.focus_set()

    def send_rdp_mouse_event(self, button, action, x, y):
        """Send mouse event to RDP server"""
        if not self.rdp_active or not self.rdp_socket:
            return

        try:
            self.rdp_socket.sendall(struct.pack('>BBHH', button, action, x, y))
        except Exception as e:
            logging.error(f"Error sending mouse event: {str(e)}")
            self.stop_rdp()

    def send_rdp_key_event(self, key, action):
        """Send keyboard event to RDP server - simplified implementation"""
        if not self.rdp_active or not self.rdp_socket:
            return

        try:
            # Simple mapping of common keys to scan codes
            # This is a basic implementation - a full implementation would map all keys
            key_map = {
                'a': 30, 'b': 48, 'c': 46, 'd': 32, 'e': 18, 'f': 33, 'g': 34, 'h': 35,
                'i': 23, 'j': 36, 'k': 37, 'l': 38, 'm': 50, 'n': 49, 'o': 24, 'p': 25,
                'q': 16, 'r': 19, 's': 31, 't': 20, 'u': 22, 'v': 47, 'w': 17, 'x': 45,
                'y': 21, 'z': 44, '1': 2, '2': 3, '3': 4, '4': 5, '5': 6, '6': 7, '7': 8,
                '8': 9, '9': 10, '0': 11, 'space': 57, 'Return': 28, 'Escape': 1,
                'BackSpace': 14, 'Tab': 15, 'Left': 75, 'Right': 77, 'Up': 72, 'Down': 80
            }

            # Convert key to lowercase for consistency
            key_lower = key.lower()

            # Get scan code
            scan_code = key_map.get(key_lower, key_map.get(key, 0))

            # Send key event
            if scan_code > 0:
                self.rdp_socket.sendall(struct.pack('>BBHH', scan_code, action, 0, 0))
        except Exception as e:
            logging.error(f"Error sending key event: {str(e)}")

    def on_tab_change(self, event=None):
        """Handle tab changes with improved RDP keyboard handling"""
        try:
            current_tab = self.notebook.select()
            prev_tab = self.active_tab
            self.active_tab = self.notebook.tab(current_tab, "text")

            print(f"Tab changed from {prev_tab} to {self.active_tab}")

            # Track if the RDP tab is active to control keyboard handling
            self.rdp_tab_active = (self.active_tab == "Remote Desktop")

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

            # Auto-load software list when entering Software tab
            if self.active_tab == "Software":
                self.auto_load_software()

            # Update the UI based on the new tab
            self.update_idletasks()

        except Exception as e:
            print(f"Error during tab change: {str(e)}")

    def auto_load_software(self):
        """Automatically load software list if a computer is selected"""
        try:
            # Only load if a computer is connected and it's not already loaded for this connection
            if self.active_connection and self.active_connection != self.software_loaded_for:
                self.update_software_status("Loading software list...")
                # Schedule software loading with a small delay to allow UI to update
                self.after(100, self.refresh_software_list)
                # Update the tracking variable
                self.software_loaded_for = self.active_connection
            elif not self.active_connection:
                self.update_software_status("Select a computer to view installed software")
        except Exception as e:
            logging.error(f"Error auto-loading software: {str(e)}")

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
