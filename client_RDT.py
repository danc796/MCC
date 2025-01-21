import tkinter as tk
from tkinter import messagebox
import socket
import struct
import threading
import time
import sys
import platform
import numpy as np
import cv2
from PIL import Image, ImageTk
import re


class RDPClient:
    """Remote Desktop Protocol Client Implementation"""

    def __init__(self):
        # Platform detection
        self.PLATFORM = self._detect_platform()

        # Configuration
        self.REFRESH_RATE = 0.05
        self.BUFFER_SIZE = 10240
        self.scale = 1.0

        # GUI setup
        self.root = tk.Tk()
        self.root.title("RDP Client")
        self.setup_gui()

        # Connection state
        self.socket = None
        self.display_window = None
        self.display_thread = None
        self.last_input_time = time.time()

        # Screen dimensions
        self.screen_width = 0
        self.screen_height = 0

        # Proxy configuration
        self.socks5_proxy = None

    def _detect_platform(self):
        """Detect the current platform"""
        if sys.platform == "win32":
            return b'win'
        elif sys.platform == "darwin":
            return b'osx'
        elif platform.system() == "Linux":
            return b'x11'
        return b''

    def setup_gui(self):
        """Set up the main GUI window"""
        # Host input
        self.host_var = tk.StringVar(value='127.0.0.1:80')
        tk.Label(self.root, text="Host:").grid(row=0, column=0, padx=10, pady=10)
        tk.Entry(self.root, textvariable=self.host_var, width=20).grid(
            row=0, column=1, padx=10, pady=10)

        # Scale slider
        tk.Label(self.root, text="Scale:").grid(row=1, column=0, padx=10, pady=10)
        scale = tk.Scale(self.root, from_=10, to=100, orient=tk.HORIZONTAL,
                         command=self._update_scale)
        scale.set(100)
        scale.grid(row=1, column=1, padx=10, pady=10)

        # Buttons
        tk.Button(self.root, text="Proxy Settings",
                  command=self._show_proxy_dialog).grid(row=2, column=0, padx=10, pady=10)
        tk.Button(self.root, text="Connect",
                  command=self._toggle_connection).grid(row=2, column=1, padx=10, pady=10)

    def _show_proxy_dialog(self):
        """Show proxy configuration dialog"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Proxy Settings")

        proxy_var = tk.StringVar(value=self.socks5_proxy or "")

        tk.Label(dialog, text="SOCKS5 Proxy:").grid(row=0, column=0, padx=10, pady=10)
        tk.Entry(dialog, textvariable=proxy_var, width=20).grid(
            row=0, column=1, padx=10, pady=10)

        def save_proxy():
            self.socks5_proxy = proxy_var.get() or None
            dialog.destroy()

        tk.Button(dialog, text="Save", command=save_proxy).grid(
            row=1, column=0, columnspan=2, pady=10)

    def _update_scale(self, value):
        """Update display scale"""
        self.scale = float(value) / 100
        if self.display_window:
            self._resize_display()

    def _resize_display(self):
        """Resize the display window"""
        if self.display_window and hasattr(self.display_window, 'canvas'):
            width = int(self.screen_width * self.scale)
            height = int(self.screen_height * self.scale)
            self.display_window.canvas.config(width=width, height=height)

    def _connect(self):
        """Establish connection to RDP server"""
        try:
            host, port = self._parse_host(self.host_var.get())

            if self.socks5_proxy:
                self._connect_via_proxy(host, port)
            else:
                self.socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
                self.socket.connect((host, port))

            self._create_display_window()

        except Exception as e:
            messagebox.showerror("Connection Error", str(e))
            self.socket = None

    def _parse_host(self, host_string):
        """Parse host string into host and port"""
        try:
            host, port = host_string.split(':')
            return host, int(port)
        except:
            raise ValueError("Invalid host format. Use host:port")

    def _connect_via_proxy(self, target_host, target_port):
        """Connect through SOCKS5 proxy"""
        try:
            proxy_host, proxy_port = self._parse_host(self.socks5_proxy)

            self.socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            self.socket.connect((proxy_host, proxy_port))

            # SOCKS5 handshake
            self.socket.sendall(struct.pack(">BB", 5, 0))
            if self.socket.recv(2)[1] != 0:
                raise ConnectionError("Proxy handshake failed")

            # Connection request
            if re.match(r'^\d+?\.\d+?\.\d+?\.\d+?$', target_host):
                # IP address
                addr = [int(x) for x in target_host.split('.')]
                req = struct.pack(">BBBB4BH", 5, 1, 0, 1, *addr, target_port)
            else:
                # Hostname
                encoded_host = target_host.encode()
                req = struct.pack(">BBBB", 5, 1, 0, 3)
                req += struct.pack(">B", len(encoded_host))
                req += encoded_host
                req += struct.pack(">H", target_port)

            self.socket.sendall(req)

            # Verify response
            if self.socket.recv(10)[1] != 0:
                raise ConnectionError("Proxy connection failed")

        except Exception as e:
            raise ConnectionError(f"Proxy connection error: {str(e)}")

    def _create_display_window(self):
        """Create the display window"""
        self.display_window = tk.Toplevel(self.root)
        self.display_window.title("Remote Display")

        # Send platform information
        self.socket.sendall(self.PLATFORM)

        # Receive initial frame
        img_type, img_data = self._receive_frame()
        img = self._process_image(img_data, is_diff=False)
        self.screen_height, self.screen_width = img.shape[:2]

        # Create canvas
        self.display_window.canvas = tk.Canvas(
            self.display_window,
            width=self.screen_width,
            height=self.screen_height
        )
        self.display_window.canvas.pack()

        # Set up event bindings
        self._setup_input_handlers(self.display_window.canvas)

        # Start display thread
        self.display_thread = threading.Thread(target=self._display_loop)
        self.display_thread.start()

    def _display_loop(self):
        """Main display loop for receiving and showing remote screen updates"""
        try:
            while True:
                # Receive and process frame
                img_type, img_data = self._receive_frame()
                img = self._process_image(img_data, is_diff=(img_type == 0))

                # Scale image if needed
                if self.scale != 1.0:
                    height, width = img.shape[:2]
                    new_size = (int(width * self.scale), int(height * self.scale))
                    img = cv2.resize(img, new_size)

                # Convert to PhotoImage for Tkinter
                img_rgb = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
                pil_img = Image.fromarray(img_rgb)
                photo_img = ImageTk.PhotoImage(image=pil_img)

                # Update canvas without recreating image object
                if not hasattr(self.display_window.canvas, 'image_id'):
                    self.display_window.canvas.image_id = self.display_window.canvas.create_image(
                        0, 0, anchor=tk.NW, image=photo_img)
                else:
                    self.display_window.canvas.itemconfig(
                        self.display_window.canvas.image_id, image=photo_img)
                self.display_window.canvas.photo = photo_img  # Prevent garbage collection

        except Exception as e:
            print(f"Display loop error: {e}")
            self._cleanup()

    def _receive_frame(self):
        """Receive a frame from the server"""
        # Receive header
        header = self._receive_exact(5)
        img_type, length = struct.unpack(">BI", header)

        # Receive image data
        img_data = b''
        while length > 0:
            chunk_size = min(self.BUFFER_SIZE, length)
            chunk = self._receive_exact(chunk_size)
            img_data += chunk
            length -= len(chunk)

        return img_type, img_data

    def _receive_exact(self, size):
        """Receive exact number of bytes from socket"""
        data = b''
        while len(data) < size:
            chunk = self.socket.recv(size - len(data))
            if not chunk:
                raise ConnectionError("Connection lost")
            data += chunk
        return data

    def _process_image(self, img_data, is_diff=False):
        """Process received image data"""
        # Decode image
        np_arr = np.frombuffer(img_data, dtype=np.uint8)
        img = cv2.imdecode(np_arr, cv2.IMREAD_COLOR)

        if is_diff and hasattr(self, 'last_image'):
            # Apply XOR difference
            img = cv2.bitwise_xor(self.last_image, img)

        self.last_image = img.copy()
        return img

    def _setup_input_handlers(self, canvas):
        """Set up input event handlers"""
        # Give canvas keyboard focus
        canvas.focus_set()

        # Mouse event bindings
        canvas.bind("<Button-1>", lambda e: self._send_mouse_event(1, 100, e.x, e.y))
        canvas.bind("<ButtonRelease-1>", lambda e: self._send_mouse_event(1, 117, e.x, e.y))
        canvas.bind("<Button-3>", lambda e: self._send_mouse_event(3, 100, e.x, e.y))
        canvas.bind("<ButtonRelease-3>", lambda e: self._send_mouse_event(3, 117, e.x, e.y))
        canvas.bind("<Motion>", self._handle_mouse_motion)

        # Platform-specific scroll bindings
        if self.PLATFORM in (b'win', b'osx'):
            canvas.bind("<MouseWheel>", self._handle_mousewheel_win_osx)
        else:  # Linux
            canvas.bind("<Button-4>", lambda e: self._send_mouse_event(2, 1, e.x, e.y))
            canvas.bind("<Button-5>", lambda e: self._send_mouse_event(2, 0, e.x, e.y))

        # Keyboard bindings
        canvas.bind("<KeyPress>", lambda e: self._send_key_event(e.keycode, 100, e.x, e.y))
        canvas.bind("<KeyRelease>", lambda e: self._send_key_event(e.keycode, 117, e.x, e.y))

    def _send_mouse_event(self, button, action, x, y):
        """Send mouse event to server"""
        if self.socket:
            scaled_x = int(x / self.scale)
            scaled_y = int(y / self.scale)
            data = struct.pack('>BBHH', button, action, scaled_x, scaled_y)
            self.socket.sendall(data)

    def _handle_mouse_motion(self, event):
        """Handle mouse motion with rate limiting"""
        current_time = time.time()
        if current_time - self.last_input_time >= self.REFRESH_RATE:
            self.last_input_time = current_time
            self._send_mouse_event(4, 0, event.x, event.y)

    def _handle_mousewheel_win_osx(self, event):
        """Handle mousewheel events for Windows and macOS"""
        delta = 1 if event.delta > 0 else 0
        self._send_mouse_event(2, delta, event.x, event.y)

    def _send_key_event(self, keycode, action, x, y):
        """Send keyboard event to server"""
        if self.socket:
            data = struct.pack('>BBHH', keycode, action, x, y)
            self.socket.sendall(data)

    def _cleanup(self):
        """Clean up resources"""
        if self.socket:
            self.socket.close()
            self.socket = None

        if self.display_window:
            self.display_window.destroy()
            self.display_window = None

    def _toggle_connection(self):
        """Toggle connection state"""
        if not self.socket:
            self._connect()
        else:
            self._cleanup()

    def run(self):
        """Start the client application"""
        self.root.mainloop()


if __name__ == '__main__':
    client = RDPClient()
    client.run()
