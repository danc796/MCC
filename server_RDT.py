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
            if key == 4:  # Mouse move
                mouse.move(x, y)

            elif key == 1:  # Left click
                if action == 100:
                    ag.mouseDown(button=ag.LEFT)
                elif action == 117:
                    ag.mouseUp(button=ag.LEFT)

            elif key == 2:  # Scroll
                scroll_amount = self.SCROLL_SENSITIVITY if action else -self.SCROLL_SENSITIVITY
                ag.scroll(scroll_amount)

            elif key == 3:  # Right click
                if action == 100:
                    ag.mouseDown(button=ag.RIGHT)
                elif action == 117:
                    ag.mouseUp(button=ag.RIGHT)

            else:  # Keyboard
                if action == 100:
                    ag.keyDown(key)
                elif action == 117:
                    ag.keyUp(key)

        except Exception as e:
            print(f"Input processing error: {e}")


if __name__ == '__main__':
    try:
        server = RDPServer()
        server.start()
    except KeyboardInterrupt:
        print("\nServer shutting down...")
    except Exception as e:
        print(f"Server error: {e}")
