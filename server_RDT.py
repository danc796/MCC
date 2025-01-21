import socket
import threading
import mss
import cv2
import numpy as np
import struct
import time
import logging
from datetime import datetime

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(f'server_log_{datetime.now().strftime("%Y%m%d_%H%M%S")}.log'),
        logging.StreamHandler()
    ]
)


class RemoteDesktopServer:
    def __init__(self, host='192.168.1.120', port=12345):
        self.host = host
        self.port = port
        self.clients = {}
        self.server_socket = None
        self.running = True
        self.lock = threading.Lock()
        self.screen_capture = None

    def initialize_server(self):
        """Initialize server socket with error handling."""
        try:
            self.server_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            self.server_socket.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            self.server_socket.bind((self.host, self.port))
            self.server_socket.listen(5)
            logging.info(f"Server initialized on {self.host}:{self.port}")
            return True
        except Exception as e:
            logging.error(f"Failed to initialize server: {e}")
            return False

    def start_screen_capture(self):
        """Initialize screen capture thread."""
        self.screen_capture = ScreenCapture(self)
        screen_thread = threading.Thread(target=self.screen_capture.capture_loop)
        screen_thread.daemon = True
        screen_thread.start()

    def accept_connections(self):
        """Accept and handle incoming client connections."""
        while self.running:
            try:
                client_socket, address = self.server_socket.accept()
                logging.info(f"New connection from {address}")

                with self.lock:
                    self.clients[address] = client_socket

                client_thread = threading.Thread(
                    target=self.handle_client,
                    args=(client_socket, address)
                )
                client_thread.daemon = True
                client_thread.start()

            except Exception as e:
                logging.error(f"Error accepting connection: {e}")
                if not self.running:
                    break

    def handle_client(self, client_socket, address):
        """Handle individual client connection."""
        try:
            while self.running:
                # Keep connection alive and handle any client-specific logic
                time.sleep(0.1)
        except Exception as e:
            logging.error(f"Error handling client {address}: {e}")
        finally:
            self.remove_client(address)

    def broadcast_frame(self, frame_data):
        """Broadcast frame data to all connected clients."""
        with self.lock:
            disconnected_clients = []

            for address, client in self.clients.items():
                try:
                    # Send frame size first
                    size = len(frame_data)
                    client.sendall(struct.pack(">L", size))

                    # Send frame data
                    client.sendall(frame_data)
                except Exception as e:
                    logging.error(f"Error sending to client {address}: {e}")
                    disconnected_clients.append(address)

            # Remove disconnected clients
            for address in disconnected_clients:
                self.remove_client(address)

    def remove_client(self, address):
        """Safely remove a client."""
        with self.lock:
            if address in self.clients:
                try:
                    self.clients[address].close()
                except Exception as e:
                    logging.error(f"Error closing client socket: {e}")
                del self.clients[address]
                logging.info(f"Client {address} removed")

    def shutdown(self):
        """Shutdown the server and cleanup resources."""
        self.running = False

        # Close all client connections
        with self.lock:
            for client in self.clients.values():
                try:
                    client.close()
                except Exception:
                    pass
            self.clients.clear()

        # Close server socket
        if self.server_socket:
            try:
                self.server_socket.close()
            except Exception:
                pass


class ScreenCapture:
    def __init__(self, server, fps=30, quality=40):
        self.server = server
        self.fps = fps
        self.quality = quality
        self.interval = 1 / fps

    def capture_loop(self):
        """Continuous screen capture loop."""
        try:
            with mss.mss() as sct:
                monitor = sct.monitors[1]  # Primary monitor

                while self.server.running:
                    start_time = time.time()

                    try:
                        # Capture screen
                        screenshot = sct.grab(monitor)
                        frame = np.array(screenshot)

                        # Encode frame
                        success, encoded_frame = cv2.imencode(
                            '.jpg',
                            frame,
                            [int(cv2.IMWRITE_JPEG_QUALITY), self.quality]
                        )

                        if success:
                            # Broadcast encoded frame
                            self.server.broadcast_frame(encoded_frame.tobytes())

                        # Maintain target FPS
                        elapsed = time.time() - start_time
                        if elapsed < self.interval:
                            time.sleep(self.interval - elapsed)

                    except Exception as e:
                        logging.error(f"Error in capture loop: {e}")
                        time.sleep(1)  # Prevent rapid-fire errors

        except Exception as e:
            logging.error(f"Critical error in screen capture: {e}")


def main():
    server = RemoteDesktopServer()

    if not server.initialize_server():
        return

    try:
        server.start_screen_capture()
        server.accept_connections()
    except KeyboardInterrupt:
        logging.info("Server shutdown requested")
    except Exception as e:
        logging.error(f"Critical server error: {e}")
    finally:
        server.shutdown()


if __name__ == "__main__":
    main()
