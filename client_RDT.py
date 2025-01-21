import socket
import tkinter as tk
from PIL import Image, ImageTk
import io
import struct
import zlib
import threading


class RemoteDesktopClient:
    def __init__(self, root, host='192.168.1.120', port=5000):
        self.root = root
        self.root.title("Remote Desktop Viewer")

        # Connection settings
        self.host = host
        self.port = port
        self.socket = None
        self.connected = False

        # Create GUI elements
        self.create_widgets()

        # Connection thread
        self.receive_thread = None

    def create_widgets(self):
        # Create the main frame
        self.main_frame = tk.Frame(self.root)
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        # Create display label
        self.display_label = tk.Label(self.main_frame)
        self.display_label.pack(padx=10, pady=10)

        # Create controls frame
        self.control_frame = tk.Frame(self.main_frame)
        self.control_frame.pack(pady=5)

        # Create connect button
        self.connect_button = tk.Button(self.control_frame, text="Connect", command=self.connect)
        self.connect_button.pack(side=tk.LEFT, padx=5)

        # Create disconnect button
        self.disconnect_button = tk.Button(self.control_frame, text="Disconnect", command=self.disconnect)
        self.disconnect_button.pack(side=tk.LEFT, padx=5)
        self.disconnect_button.config(state=tk.DISABLED)

    def connect(self):
        try:
            self.socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            self.socket.connect((self.host, self.port))
            self.connected = True

            # Update button states
            self.connect_button.config(state=tk.DISABLED)
            self.disconnect_button.config(state=tk.NORMAL)

            # Start receiving screen updates
            self.receive_thread = threading.Thread(target=self.receive_screen_updates)
            self.receive_thread.daemon = True
            self.receive_thread.start()

        except Exception as e:
            print(f"Connection error: {e}")
            self.connected = False

    def receive_screen_updates(self):
        try:
            while self.connected:
                # Receive size of the incoming data
                size_data = self.socket.recv(4)
                size = struct.unpack('>I', size_data)[0]

                # Receive the compressed image data
                received_data = b""
                while len(received_data) < size:
                    chunk = self.socket.recv(min(size - len(received_data), 8192))
                    if not chunk:
                        raise Exception("Connection lost")
                    received_data += chunk

                # Decompress the data
                img_data = zlib.decompress(received_data)

                # Convert to image
                image = Image.open(io.BytesIO(img_data))
                photo = ImageTk.PhotoImage(image)

                # Update the display
                self.display_label.config(image=photo)
                self.display_label.image = photo

        except Exception as e:
            print(f"Error receiving screen updates: {e}")
            self.disconnect()

    def disconnect(self):
        self.connected = False
        if self.socket:
            self.socket.close()

        # Update button states
        self.connect_button.config(state=tk.NORMAL)
        self.disconnect_button.config(state=tk.DISABLED)

        # Clear the display
        self.display_label.config(image='')


def main():
    root = tk.Tk()
    client = RemoteDesktopClient(root)
    root.mainloop()


if __name__ == "__main__":
    main()
