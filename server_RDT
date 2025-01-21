import socket
import threading
import zlib
from PIL import ImageGrab, Image
import io
import struct


class RemoteDesktopServer:
    def __init__(self, host='0.0.0.0', port=5000):
        self.host = host
        self.port = port
        self.server_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        self.running = False
        self.clients = []

    def start(self):
        self.server_socket.bind((self.host, self.port))
        self.server_socket.listen(5)
        self.running = True
        print(f"Server listening on {self.host}:{self.port}")

        # Accept client connections in a separate thread
        accept_thread = threading.Thread(target=self.accept_connections)
        accept_thread.daemon = True
        accept_thread.start()

    def accept_connections(self):
        while self.running:
            try:
                client_socket, address = self.server_socket.accept()
                print(f"Client connected from {address}")
                self.clients.append(client_socket)

                # Start a new thread to handle this client
                client_thread = threading.Thread(target=self.handle_client, args=(client_socket,))
                client_thread.daemon = True
                client_thread.start()
            except:
                break

    def handle_client(self, client_socket):
        try:
            while self.running:
                # Capture screen
                screenshot = ImageGrab.grab()

                # Resize to reduce data size (adjust as needed)
                screenshot = screenshot.resize((1280, 1280), Image.Resampling.LANCZOS)

                # Convert to bytes
                img_byte_arr = io.BytesIO()
                screenshot.save(img_byte_arr, format='JPEG', quality=50)
                img_byte_arr = img_byte_arr.getvalue()

                # Compress
                compressed_data = zlib.compress(img_byte_arr, level=6)

                # Send size first
                size = len(compressed_data)
                client_socket.send(struct.pack('>I', size))

                # Send data
                client_socket.send(compressed_data)

        except Exception as e:
            print(f"Error handling client: {e}")
        finally:
            if client_socket in self.clients:
                self.clients.remove(client_socket)
            client_socket.close()

    def stop(self):
        self.running = False
        for client in self.clients:
            client.close()
        self.server_socket.close()


def main():
    server = RemoteDesktopServer()
    try:
        server.start()
        input("Press Enter to stop the server...\n")
    except KeyboardInterrupt:
        pass
    finally:
        server.stop()


if __name__ == "__main__":
    main()
