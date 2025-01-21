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
import keyboard


class RDPClient:
    def __init__(self):
        # Basic configuration
        self.REFRESH_RATE = 0.05
        self.BUFFER_SIZE = 10240
        self.scale = 1.0
        self.platform = self._detect_platform()

        # Setup main window
        self.root = tk.Tk()
        self.root.title("RDP Client")
        self.setup_gui()

        # Connection state
        self.socket = None
        self.display_window = None
        self.display_thread = None
        self.keyboard_thread = None
        self.keyboard_active = False
        self.last_input_time = time.time()

        # Screen dimensions
        self.screen_width = 0
        self.screen_height = 0

    def _detect_platform(self):
        if sys.platform == "win32":
            return b'win'
        elif sys.platform == "darwin":
            return b'osx'
        elif platform.system() == "Linux":
            return b'x11'
        return b''

    def setup_gui(self):
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

        # Connect button
        tk.Button(self.root, text="Connect/Disconnect",
                  command=self.toggle_connection).grid(row=2, column=0, columnspan=2, padx=10, pady=10)

    def _update_scale(self, value):
        self.scale = float(value) / 100
        if self.display_window:
            self._resize_display()

    def _resize_display(self):
        if self.display_window and hasattr(self.display_window, 'canvas'):
            width = int(self.screen_width * self.scale)
            height = int(self.screen_height * self.scale)
            self.display_window.canvas.config(width=width, height=height)

    def _connect(self):
        try:
            host, port = self.host_var.get().split(':')
            self.socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            self.socket.connect((host, int(port)))
            self._create_display_window()

        except Exception as e:
            messagebox.showerror("Connection Error", str(e))
            self.socket = None

    def _create_display_window(self):
        self.display_window = tk.Toplevel(self.root)
        self.display_window.title("Remote Display")

        # Send platform information
        self.socket.sendall(self.platform)

        # Get initial frame
        img_type, img_data = self._receive_frame()
        img = self._process_image(img_data, is_diff=False)
        self.screen_height, self.screen_width = img.shape[:2]

        # Create display canvas
        self.display_window.canvas = tk.Canvas(
            self.display_window,
            width=self.screen_width,
            height=self.screen_height
        )
        self.display_window.canvas.pack()

        # Start input and display handlers
        self._setup_input_handlers()
        self.display_thread = threading.Thread(target=self._display_loop)
        self.display_thread.start()

    def _setup_input_handlers(self):
        canvas = self.display_window.canvas
        canvas.focus_set()

        # Updated mouse handlers with new identifiers
        canvas.bind("<Button-1>", lambda e: self._send_mouse_event(self.MOUSE_LEFT, 100, e.x, e.y))
        canvas.bind("<ButtonRelease-1>", lambda e: self._send_mouse_event(self.MOUSE_LEFT, 117, e.x, e.y))
        canvas.bind("<Button-3>", lambda e: self._send_mouse_event(self.MOUSE_RIGHT, 100, e.x, e.y))
        canvas.bind("<ButtonRelease-3>", lambda e: self._send_mouse_event(self.MOUSE_RIGHT, 117, e.x, e.y))
        canvas.bind("<Motion>", self._handle_mouse_motion)

        if self.platform in (b'win', b'osx'):
            canvas.bind("<MouseWheel>", self._handle_mousewheel)
        else:
            canvas.bind("<Button-4>", lambda e: self._send_mouse_event(self.MOUSE_SCROLL, 1, e.x, e.y))
            canvas.bind("<Button-5>", lambda e: self._send_mouse_event(self.MOUSE_SCROLL, 0, e.x, e.y))

        self.keyboard_active = True
        self.keyboard_thread = threading.Thread(target=self._keyboard_loop)
        self.keyboard_thread.daemon = True
        self.keyboard_thread.start()

    def _keyboard_loop(self):
        """Dedicated keyboard monitoring loop with mouse button handling"""
        while self.keyboard_active:
            try:
                event = keyboard.read_event(suppress=True)
                if event.event_type in ('down', 'up'):
                    # Check if it's a mouse button event
                    if event.name in ('mouse_left', 'mouse_right', 'mouse_middle'):
                        # Handle mouse events as before
                        button = 1 if event.name == 'mouse_left' else 3 if event.name == 'mouse_right' else 2
                        action = 100 if event.event_type == 'down' else 117
                        self._send_mouse_event(button, action, 0, 0)
                    else:
                        # Handle keyboard events
                        self._send_keyboard_event(event)
            except Exception as e:
                print(f"Keyboard error: {e}")
                break

    def _send_keyboard_event(self, event):
        """Send keyboard event to server with numeric key handling"""
        if not self.socket:
            return

        try:
            action = 100 if event.event_type == 'down' else 117
            scan_code = event.scan_code

            # Send keyboard event to server
            data = struct.pack('>BBHH', scan_code, action, 0, 0)
            self.socket.sendall(data)
        except Exception as e:
            print(f"Failed to send keyboard event: {e}")
            self._cleanup()

    def _send_mouse_event(self, button, action, x, y):
        """Send mouse event to server"""
        if self.socket:
            scaled_x = int(x / self.scale)
            scaled_y = int(y / self.scale)
            try:
                self.socket.sendall(struct.pack('>BBHH', button, action, scaled_x, scaled_y))
            except:
                self._cleanup()

    def _handle_mouse_motion(self, event):
        current_time = time.time()
        if current_time - self.last_input_time >= self.REFRESH_RATE:
            self.last_input_time = current_time
            self._send_mouse_event(self.MOUSE_MOVE, 0, event.x, event.y)

    def _handle_mousewheel(self, event):
        delta = 1 if event.delta > 0 else 0
        self._send_mouse_event(self.MOUSE_SCROLL, delta, event.x, event.y)

    def _receive_frame(self):
        """Receive a frame from the server"""
        # Get header
        header = self._receive_exact(5)
        img_type, length = struct.unpack(">BI", header)

        # Get image data
        img_data = b''
        while length > 0:
            chunk_size = min(self.BUFFER_SIZE, length)
            chunk = self._receive_exact(chunk_size)
            img_data += chunk
            length -= len(chunk)

        return img_type, img_data

    def _receive_exact(self, size):
        """Receive exact number of bytes"""
        data = b''
        while len(data) < size:
            chunk = self.socket.recv(size - len(data))
            if not chunk:
                raise ConnectionError("Connection lost")
            data += chunk
        return data

    def _process_image(self, img_data, is_diff=False):
        """Process received image data with direct color handling"""
        np_arr = np.frombuffer(img_data, dtype=np.uint8)
        img = cv2.imdecode(np_arr, cv2.IMREAD_COLOR)

        if is_diff and hasattr(self, 'last_image'):
            img = cv2.bitwise_xor(self.last_image, img)

        self.last_image = img.copy()
        return img

    def _display_loop(self):
        """Main display loop with stable color handling"""
        try:
            while True:
                img_type, img_data = self._receive_frame()
                img = self._process_image(img_data, is_diff=(img_type == 0))

                if self.scale != 1.0:
                    height, width = img.shape[:2]
                    new_size = (int(width * self.scale), int(height * self.scale))
                    img = cv2.resize(img, new_size, interpolation=cv2.INTER_LINEAR)

                # Simple BGR to RGB conversion without additional processing
                img_rgb = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)

                pil_img = Image.fromarray(img_rgb)
                photo_img = ImageTk.PhotoImage(image=pil_img)

                if not hasattr(self.display_window.canvas, 'image_id'):
                    self.display_window.canvas.image_id = self.display_window.canvas.create_image(
                        0, 0, anchor=tk.NW, image=photo_img)
                else:
                    self.display_window.canvas.itemconfig(
                        self.display_window.canvas.image_id, image=photo_img)
                self.display_window.canvas.photo = photo_img

        except Exception as e:
            print(f"Display error: {e}")
            self._cleanup()
        finally:
            self._cleanup()

    def _cleanup(self):
        """Enhanced cleanup with keyboard handling"""
        self.keyboard_active = False
        if self.keyboard_thread:
            self.keyboard_thread.join(timeout=1.0)

        if self.socket:
            self.socket.close()
            self.socket = None

        if self.display_window:
            self.display_window.destroy()
            self.display_window = None

    def toggle_connection(self):
        """Toggle connection state"""
        if not self.socket:
            self._connect()
        else:
            self._cleanup()

    def run(self):
        """Start the client application"""
        self.root.mainloop()

    MOUSE_LEFT = 201
    MOUSE_SCROLL = 202
    MOUSE_RIGHT = 203
    MOUSE_MOVE = 204


if __name__ == '__main__':
    client = RDPClient()
    client.run()
