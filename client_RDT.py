import socket
import pygame
import io
import struct
import logging
import sys
import time
from PIL import Image
from datetime import datetime

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(f'client_log_{datetime.now().strftime("%Y%m%d_%H%M%S")}.log'),
        logging.StreamHandler()
    ]
)


class RemoteDesktopClient:
    def __init__(self, host='192.168.1.120', port=12345):
        self.host = host
        self.port = port
        self.socket = None
        self.screen = None
        self.running = True
        self.reconnect_delay = 5
        self.max_reconnect_delay = 60

        # Initialize Pygame
        pygame.init()

    def connect(self):
        """Connect to server with automatic reconnection."""
        while self.running:
            try:
                self.socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
                self.socket.connect((self.host, self.port))
                logging.info("Connected to server successfully")
                self.reconnect_delay = 5  # Reset delay on successful connection
                return True
            except Exception as e:
                logging.error(f"Connection failed: {e}")
                self.cleanup_socket()

                if not self.running:
                    break

                logging.info(f"Retrying in {self.reconnect_delay} seconds...")
                time.sleep(self.reconnect_delay)
                self.reconnect_delay = min(self.reconnect_delay * 2, self.max_reconnect_delay)

        return False

    def recvall(self, size):
        """Receive the exact amount of data with timeout."""
        try:
            data = bytearray()
            while len(data) < size and self.running:
                try:
                    chunk = self.socket.recv(min(4096, size - len(data)))
                    if not chunk:
                        raise ConnectionError("Connection closed by server")
                    data.extend(chunk)
                except socket.timeout:
                    continue
                except Exception as e:
                    raise ConnectionError(f"Error receiving data: {e}")
            return bytes(data)
        except Exception as e:
            logging.error(f"Error in recvall: {e}")
            raise

    def handle_events(self):
        """Handle Pygame events."""
        for event in pygame.event.get():
            if event.type == pygame.QUIT:
                self.running = False
                return False
            # Add additional event handling here (keyboard, mouse, etc.)
        return True

    def receive_frame(self):
        """Receive and process a single frame."""
        try:
            # Receive frame size
            size_data = self.recvall(4)
            if not size_data:
                return False

            frame_size = struct.unpack(">L", size_data)[0]

            # Receive frame data
            frame_data = self.recvall(frame_size)
            if not frame_data:
                return False

            # Convert to PIL Image
            image = Image.open(io.BytesIO(frame_data))

            # Initialize or update Pygame display if needed
            if self.screen is None:
                self.screen = pygame.display.set_mode(image.size)
                pygame.display.set_caption("Remote Desktop Viewer")

            # Convert image to RGB mode for Pygame compatibility
            if image.mode != 'RGB':
                image = image.convert('RGB')

            # Convert to Pygame surface and display
            frame = pygame.image.fromstring(image.tobytes(), image.size, 'RGB')
            self.screen.blit(frame, (0, 0))
            pygame.display.flip()

            return True

        except ConnectionError as e:
            logging.error(f"Connection error: {e}")
            return False
        except Exception as e:
            logging.error(f"Error processing frame: {e}")
            return False

    def main_loop(self):
        """Main client loop with error handling and reconnection."""
        while self.running:
            try:
                if not self.connect():
                    continue

                self.socket.settimeout(1.0)  # Set socket timeout

                while self.running:
                    if not self.handle_events():
                        break

                    if not self.receive_frame():
                        break

            except Exception as e:
                logging.error(f"Error in main loop: {e}")
                if self.running:
                    self.cleanup_socket()
                    continue

        self.cleanup()

    def cleanup_socket(self):
        """Clean up socket connection."""
        if self.socket:
            try:
                self.socket.close()
            except Exception as e:
                logging.error(f"Error closing socket: {e}")
            self.socket = None

    def cleanup(self):
        """Clean up all resources."""
        self.running = False
        self.cleanup_socket()
        pygame.quit()
        logging.info("Client shutdown complete")


def main():
    try:
        client = RemoteDesktopClient()
        client.main_loop()
    except KeyboardInterrupt:
        logging.info("Client shutdown requested")
    except Exception as e:
        logging.error(f"Critical client error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
