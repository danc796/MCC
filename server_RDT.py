import struct
import socket
import numpy as np
import cv2
import threading
import time
import pyautogui as ag
import mouse
from PIL import ImageGrab


class RDPServer:
    def __init__(self, host='0.0.0.0', port=80):
        # Configuration
        self.REFRESH_RATE = 0.05
        self.SCROLL_SENSITIVITY = 5
        self.IMAGE_QUALITY = 95
        self.BUFFER_SIZE = 1024

        # Server setup
        self.host = (host, port)
        self.socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        self.socket.bind(self.host)
        self.socket.listen(1)

        # Image state
        self.last_image = None
        self.lock = threading.Lock()

        self.shift_pressed = False

    def start(self):
        """Start the RDP server and listen for connections"""
        print(f"Server started on {self.host[0]}:{self.host[1]}")
        while True:
            conn, addr = self.socket.accept()
            print(f"New connection from {addr}")

            # Start display and input threads for the client
            threading.Thread(target=self.handle_display, args=(conn,)).start()
            threading.Thread(target=self.handle_input, args=(conn,)).start()

    def handle_display(self, conn):
        """Handle screen capture and transmission"""
        try:
            # Initial screen capture and send
            initial_image = np.array(ImageGrab.grab())
            initial_image = cv2.cvtColor(initial_image, cv2.COLOR_RGB2BGR)
            _, image_bytes = cv2.imencode('.jpg', initial_image,
                                          [cv2.IMWRITE_JPEG_QUALITY, self.IMAGE_QUALITY])

            # Send initial frame
            header = struct.pack(">BI", 1, len(image_bytes))
            conn.sendall(header)
            conn.sendall(image_bytes)

            self.last_image = initial_image

            # Continuous screen update loop
            while True:
                time.sleep(self.REFRESH_RATE)

                # Capture new screen
                screen = np.array(ImageGrab.grab())
                screen = cv2.cvtColor(screen, cv2.COLOR_RGB2BGR)

                # Check for changes
                if np.array_equal(screen, self.last_image):
                    continue

                # Encode and send new frame
                _, frame_data = cv2.imencode('.jpg', screen,
                                             [cv2.IMWRITE_JPEG_QUALITY, self.IMAGE_QUALITY])

                header = struct.pack(">BI", 1, len(frame_data))
                conn.sendall(header)
                conn.sendall(frame_data)

                self.last_image = screen

        except Exception as e:
            print(f"Display handling error: {e}")
            conn.close()

    def handle_input(self, conn):
        """Handle input events from client"""
        try:
            # Get client platform info
            platform = conn.recv(3)
            print(f"Client platform: {platform.decode()}")

            # Input event loop
            while True:
                event_data = conn.recv(6)
                if not event_data or len(event_data) != 6:
                    break

                key, action, x, y = struct.unpack('>BBHH', event_data)
                self.process_input(key, action, x, y)

        except Exception as e:
            print(f"Input handling error: {e}")
            conn.close()

    def process_input(self, key, action, x, y):
        """Process individual input events"""
        try:
            # Update shift key state
            if key in (42, 54):  # Left or right shift
                self.shift_pressed = (action == 100)  # 100 for keydown, 117 for keyup
                return

            # Handle keyboard input
            if key < 200:  # Not a mouse event
                try:
                    if action == 100:  # Key down
                        ag.keyDown(self._scan_to_key(key, self.shift_pressed))
                    elif action == 117:  # Key up
                        ag.keyUp(self._scan_to_key(key, self.shift_pressed))
                except KeyError:
                    print(f"Unrecognized scan code: {key}")
            else:
                # Handle mouse events as before
                if key == self.MOUSE_MOVE:
                    mouse.move(x, y)
                elif key == self.MOUSE_LEFT:
                    if action == 100:
                        ag.mouseDown(button=ag.LEFT)
                    elif action == 117:
                        ag.mouseUp(button=ag.LEFT)
                elif key == self.MOUSE_SCROLL:
                    scroll_amount = self.SCROLL_SENSITIVITY if action else -self.SCROLL_SENSITIVITY
                    ag.scroll(scroll_amount)
                elif key == self.MOUSE_RIGHT:
                    if action == 100:
                        ag.mouseDown(button=ag.RIGHT)
                    elif action == 117:
                        ag.mouseUp(button=ag.RIGHT)

        except Exception as e:
            print(f"Input processing error: {e}")

    def _scan_to_key(self, scan_code, shift_pressed):
        """Convert scan code to PyAutoGUI key name, considering shift state"""
        scan_map = {
            30: 'a', 48: 'b', 46: 'c', 32: 'd', 18: 'e', 33: 'f', 34: 'g', 35: 'h',
            23: 'i', 36: 'j', 37: 'k', 38: 'l', 50: 'm', 49: 'n', 24: 'o', 25: 'p',
            16: 'q', 19: 'r', 31: 's', 20: 't', 22: 'u', 47: 'v', 17: 'w', 45: 'x',
            21: 'y', 44: 'z', 2: '1', 3: '2', 4: '3', 5: '4', 6: '5', 7: '6', 8: '7',
            9: '8', 10: '9', 11: '0', 26: '[', 27: ']', 43: '\\', 39: ';', 40: "'",
            41: '`', 51: ',', 52: '.', 53: '/', 12: '-', 13: '=', 28: 'enter',
            1: 'esc', 14: 'backspace', 15: 'tab', 57: 'space', 42: 'shift',
            54: 'rshift', 29: 'ctrl', 56: 'alt', 72: 'up', 80: 'down', 75: 'left',
            77: 'right', 59: 'f1', 60: 'f2', 61: 'f3', 62: 'f4', 63: 'f5', 64: 'f6',
            65: 'f7', 66: 'f8', 67: 'f9', 68: 'f10', 87: 'f11', 88: 'f12', 83: 'delete',
            71: 'home', 79: 'end', 81: 'pagedown', 73: 'pageup', 55: '*', 74: '-',
            78: '+'
        }

        # Define shifted keys dynamically
        shifted_map = {
            '1': '!', '2': '@', '3': '#', '4': '$', '5': '%', '6': '^', '7': '&',
            '8': '*', '9': '(', '0': ')', '[': '{', ']': '}', '\\': '|', ';': ':',
            "'": '"', '`': '~', ',': '<', '.': '>', '/': '?', '-': '_', '=': '+'
        }

        # Get the base key
        key = scan_map.get(scan_code)
        if not key:
            return None

        # Return shifted key if Shift is pressed and a shifted variant exists
        if shift_pressed and key in shifted_map:
            return shifted_map[key]

        return key

    MOUSE_LEFT = 201
    MOUSE_SCROLL = 202
    MOUSE_RIGHT = 203
    MOUSE_MOVE = 204


if __name__ == '__main__':
    try:
        server = RDPServer()
        server.start()
    except KeyboardInterrupt:
        print("\nServer shutting down...")
    except Exception as e:
        print(f"Server error: {e}")
