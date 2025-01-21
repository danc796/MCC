import socket
import struct
import threading
import time
from PIL import ImageGrab
import cv2
import numpy as np
import pyautogui as ag
import mouse


class RDPServer:
    """Remote Desktop Protocol Server Implementation"""

    def __init__(self, host='0.0.0.0', port=80):
        # Configuration
        self.REFRESH_RATE = 0.05
        self.SCROLL_SENSITIVITY = 5
        self.IMAGE_QUALITY = 50
        self.BUFFER_SIZE = 1024

        # Server setup
        self.host = (host, port)
        self.socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        self.socket.bind(self.host)
        self.socket.listen(1)

        # Screen capture state
        self.current_image = None
        self.current_image_bytes = None
        self.lock = threading.Lock()

        # Keyboard mapping for different platforms
        self.KEYBOARD_MAPPING = {
            b'win': self._get_windows_keymap(),
            b'osx': self._get_osx_keymap(),
            b'x11': self._get_x11_keymap()
        }

    def start(self):
        """Start the RDP server"""
        print(f"Server started on {self.host[0]}:{self.host[1]}")
        while True:
            conn, addr = self.socket.accept()
            print(f"New connection from {addr}")
            threading.Thread(target=self._handle_display, args=(conn,)).start()
            threading.Thread(target=self._handle_input, args=(conn,)).start()

    def _handle_display(self, conn):
        """Handle screen capture and transmission"""
        try:
            # Send initial screen capture
            self._capture_and_send_screen(conn, force_full=True)

            # Continuous screen update loop
            while True:
                time.sleep(self.REFRESH_RATE)
                self._capture_and_send_screen(conn)
        except Exception as e:
            print(f"Display handling error: {e}")
            conn.close()

    def _capture_and_send_screen(self, conn, force_full=False):
        """Capture screen and send to client"""
        with self.lock:
            # Capture new screen
            screen = np.array(ImageGrab.grab())
            _, new_bytes = cv2.imencode(".jpg", screen,
                                        [cv2.IMWRITE_JPEG_QUALITY, self.IMAGE_QUALITY])

            if self.current_image is None or force_full:
                # Send full image
                self._send_frame(conn, new_bytes, is_diff=False)
                self.current_image = cv2.imdecode(np.array(new_bytes), cv2.IMREAD_COLOR)
                self.current_image_bytes = new_bytes
            else:
                # Calculate and send difference if significant
                new_image = cv2.imdecode(np.array(new_bytes), cv2.IMREAD_COLOR)
                diff = cv2.absdiff(new_image, self.current_image)

                if np.any(diff):
                    _, diff_bytes = cv2.imencode(".png", diff)
                    if len(diff_bytes) < len(new_bytes):
                        self._send_frame(conn, diff_bytes, is_diff=True)
                    else:
                        self._send_frame(conn, new_bytes, is_diff=False)

                    self.current_image = new_image
                    self.current_image_bytes = new_bytes

    def _send_frame(self, conn, image_bytes, is_diff):
        """Send an image frame to the client"""
        header = struct.pack(">BI", 0 if is_diff else 1, len(image_bytes))
        conn.sendall(header)
        conn.sendall(image_bytes)

    def _handle_input(self, conn):
        """Handle input events from client"""
        try:
            # Get client platform
            platform = conn.recv(3)
            print(f"Client platform: {platform.decode()}")
            keymap = self.KEYBOARD_MAPPING.get(platform, {})

            # Input event loop
            while True:
                event_data = conn.recv(6)
                if not event_data or len(event_data) != 6:
                    break

                key, action, x, y = struct.unpack('>BBHH', event_data)
                self._process_input_event(key, action, x, y, keymap)
        except Exception as e:
            print(f"Input handling error: {e}")
            conn.close()

    def _process_input_event(self, key, action, x, y, keymap):
        """Process individual input events"""
        if key == 4:  # Mouse move
            mouse.move(x, y)
        elif key == 1:  # Left click
            if action == 100:
                ag.mouseDown(button=ag.LEFT)
            elif action == 117:
                ag.mouseUp(button=ag.LEFT)
        elif key == 2:  # Scroll
            ag.scroll(self.SCROLL_SENSITIVITY if action else -self.SCROLL_SENSITIVITY)
        elif key == 3:  # Right click
            if action == 100:
                ag.mouseDown(button=ag.RIGHT)
            elif action == 117:
                ag.mouseUp(button=ag.RIGHT)
        else:  # Keyboard
            key_name = keymap.get(key)
            if key_name:
                if action == 100:
                    ag.keyDown(key_name)
                elif action == 117:
                    ag.keyUp(key_name)

    @staticmethod
    def _get_windows_keymap():
        """Get Windows keyboard mapping"""
        return {
            0x08: 'backspace',
            0x09: 'tab',
            0x0d: 'enter',
            0x10: 'shift',
            0x11: 'ctrl',
            0x12: 'alt',
            # Add more Windows virtual key codes as needed
        }

    @staticmethod
    def _get_osx_keymap():
        """Get macOS keyboard mapping"""
        return {
            51: 'backspace',
            48: 'tab',
            36: 'enter',
            56: 'shift',
            59: 'ctrl',
            58: 'alt',
            # Add more macOS key codes as needed
        }

    @staticmethod
    def _get_x11_keymap():
        """Get X11 keyboard mapping"""
        return {
            22: 'backspace',
            23: 'tab',
            36: 'enter',
            50: 'shift',
            37: 'ctrl',
            64: 'alt',
            # Add more X11 key codes as needed
        }


if __name__ == '__main__':
    server = RDPServer()
    server.start()
