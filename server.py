import socket
import threading
import json
import psutil
import platform
import subprocess
import os
from datetime import datetime
import winreg
import logging
from cryptography.fernet import Fernet
import sys
import time
import numpy as np
import cv2
import pyautogui as ag
import mouse
from PIL import ImageGrab
import struct


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
        self.socket.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)  # Add this line to allow socket reuse
        self.socket.bind(self.host)
        self.socket.listen(1)

        # Set a timeout to make the server stoppable
        self.socket.settimeout(1.0)

        # Image state
        self.last_image = None
        self.lock = threading.Lock()
        self.shift_pressed = False

        # Control flags
        self.running = True
        self.active_connections = []
        self.threads = []

    def start(self):
        """Start the RDP server and listen for connections with improved error handling"""
        print(f"Server started on {self.host[0]}:{self.host[1]}")

        while self.running:
            try:
                conn, addr = self.socket.accept()
                print(f"New connection from {addr}")

                # Store connection reference
                self.active_connections.append(conn)

                # Start display and input threads for the client
                display_thread = threading.Thread(target=self.handle_display, args=(conn,))
                input_thread = threading.Thread(target=self.handle_input, args=(conn,))

                display_thread.daemon = True
                input_thread.daemon = True

                self.threads.append(display_thread)
                self.threads.append(input_thread)

                display_thread.start()
                input_thread.start()

            except socket.timeout:
                # This is normal - allows checking the running flag
                continue
            except Exception as e:
                if self.running:  # Only log if we weren't intentionally stopped
                    print(f"RDP server connection error: {e}")
                break

    def stop(self):
        """Stop the RDP server and clean up connections"""
        print("Stopping RDP server...")
        self.running = False

        # Close all active connections
        for conn in self.active_connections:
            try:
                conn.shutdown(socket.SHUT_RDWR)
                conn.close()
            except:
                pass

        self.active_connections = []

        # Close main socket
        try:
            self.socket.close()
        except:
            pass

        print("RDP server stopped")

    def handle_display(self, conn):
        """Handle screen capture and transmission with improved error handling"""
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
            while self.running and conn in self.active_connections:
                time.sleep(self.REFRESH_RATE)

                # Capture new screen
                try:
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
                except:
                    # Handle screen capture errors - just continue to next frame
                    continue

        except Exception as e:
            print(f"Display handling error: {e}")
        finally:
            # Clean up the connection if it's still in our list
            if conn in self.active_connections:
                try:
                    conn.close()
                    self.active_connections.remove(conn)
                except:
                    pass

    def handle_input(self, conn):
        """Handle input events from client with improved error handling"""
        try:
            # Get client platform info
            platform = conn.recv(3)
            print(f"Client platform: {platform.decode()}")

            # Input event loop
            while self.running and conn in self.active_connections:
                try:
                    event_data = conn.recv(6)
                    if not event_data or len(event_data) != 6:
                        break

                    key, action, x, y = struct.unpack('>BBHH', event_data)
                    self.process_input(key, action, x, y)
                except socket.timeout:
                    continue
                except Exception as e:
                    print(f"Input reception error: {e}")
                    break

        except Exception as e:
            print(f"Input handling error: {e}")
        finally:
            # Clean up the connection if it's still in our list
            if conn in self.active_connections:
                try:
                    conn.close()
                    self.active_connections.remove(conn)
                except:
                    pass

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


class MCCServer:
    def __init__(self, host='0.0.0.0', port=5000):
        # Setup logging first
        logging.basicConfig(
            level=logging.DEBUG,
            format='%(asctime)s - %(levelname)s - [%(filename)s:%(lineno)d] - %(message)s',
            handlers=[
                logging.FileHandler('server.log'),
                logging.StreamHandler(sys.stdout)
            ]
        )

        logging.info("Starting MCC Server...")

        self.host = host
        self.port = port
        self.server_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        self.clients = {}
        self.encryption_key = Fernet.generate_key()
        self.cipher_suite = Fernet(self.encryption_key)
        self.running = True

        self.rdp_server = None
        self.rdp_thread = None

    def start(self):
        try:
            self.server_socket.bind((self.host, self.port))
            self.server_socket.listen(2)
            self.server_socket.settimeout(1.0)

            while self.running:
                try:
                    client_socket, address = self.server_socket.accept()
                    client_handler = threading.Thread(target=self.handle_client, args=(client_socket, address))
                    client_handler.daemon = True
                    client_handler.start()
                except socket.timeout:
                    continue
                except Exception as e:
                    logging.error(f"Error accepting connection: {str(e)}")
                    if self.running:
                        logging.exception("Server error")

        except Exception as e:
            logging.error(f"Server error: {str(e)}")
        finally:
            self.stop()

    def handle_client(self, client_socket, address):
        try:
            client_socket.settimeout(1.0)
            logging.info(f"New connection from {address}")

            client_socket.send(self.encryption_key)

            self.clients[address] = {
                'socket': client_socket,
                'last_seen': datetime.now(),
                'system_info': self.get_system_info()
            }

            while self.running:
                try:
                    encrypted_data = client_socket.recv(4096)
                    if not encrypted_data:
                        break

                    data = self.cipher_suite.decrypt(encrypted_data).decode()
                    command = json.loads(data)
                    response = self.process_command(command)

                    encrypted_response = self.cipher_suite.encrypt(json.dumps(response).encode())
                    client_socket.send(encrypted_response)

                except socket.timeout:
                    continue
                except Exception as e:
                    logging.error(f"Error handling client {address}: {str(e)}")
                    break

        except Exception as e:
            logging.error(f"Client handler error for {address}: {str(e)}")
        finally:
            self.clients.pop(address, None)
            try:
                client_socket.close()
            except:
                pass
            logging.info(f"Connection closed from {address}")

    def get_system_info(self):
        """Gather system information"""
        return {
            'hostname': platform.node(),
            'os': platform.system(),
            'os_version': platform.version(),
            'cpu_count': psutil.cpu_count(),
            'total_memory': psutil.virtual_memory().total,
            'disk_partitions': [partition.mountpoint for partition in psutil.disk_partitions()]
        }

    def process_command(self, command):
        cmd_type = command.get('type')
        cmd_data = command.get('data', {})

        command_handlers = {
            'system_info': self.handle_system_info,
            'hardware_monitor': self.handle_hardware_monitor,
            'software_inventory': self.handle_software_inventory,
            'power_management': self.handle_power_management,
            'execute_command': self.handle_command_execution,
            'network_monitor': self.handle_network_monitor,
            'start_rdp': self.handle_start_rdp,
            'stop_rdp': self.handle_stop_rdp
        }

        handler = command_handlers.get(cmd_type)
        if handler:
            return handler(cmd_data)
        else:
            return {'status': 'error', 'message': 'Unknown command'}

    def handle_system_info(self, data):
        """Return system information"""
        return {
            'status': 'success',
            'data': self.get_system_info()
        }

    def handle_hardware_monitor(self, data):
        """Monitor hardware metrics with improved drive handling"""
        disk_usage = {}

        # Safely collect disk usage information
        for partition in psutil.disk_partitions(all=False):
            try:
                # Only check fixed drives and skip removable drives
                if 'fixed' in partition.opts or partition.fstype == 'NTFS':
                    usage = psutil.disk_usage(partition.mountpoint)
                    disk_usage[partition.mountpoint] = dict(usage._asdict())
            except Exception as e:
                logging.warning(f"Could not access drive {partition.mountpoint}: {str(e)}")
                continue

        return {
            'status': 'success',
            'data': {
                'cpu_percent': psutil.cpu_percent(interval=1),
                'memory_usage': dict(psutil.virtual_memory()._asdict()),
                'disk_usage': disk_usage,
                'network_io': dict(psutil.net_io_counters()._asdict())
            }
        }

    def handle_software_inventory(self, data):
        """Get installed software inventory with improved registry handling"""
        print("\n=== Starting Software Inventory Scan ===")
        software_list = []

        try:
            if platform.system() != 'Windows':
                return {
                    'status': 'error',
                    'message': 'Not a Windows system'
                }

            # Registry paths to check
            keys_to_check = [
                (winreg.HKEY_LOCAL_MACHINE, r'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'),
                (winreg.HKEY_LOCAL_MACHINE, r'SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'),
                (winreg.HKEY_CURRENT_USER, r'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall')
            ]

            seen_programs = set()

            for reg_root, key_path in keys_to_check:
                reg_key = None
                try:
                    reg_key = winreg.OpenKey(reg_root, key_path)
                    subkey_count, _, _ = winreg.QueryInfoKey(reg_key)

                    for i in range(subkey_count):
                        subkey = None
                        try:
                            subkey_name = winreg.EnumKey(reg_key, i)
                            subkey = winreg.OpenKey(reg_key, subkey_name)

                            try:
                                name = winreg.QueryValueEx(subkey, "DisplayName")[0].strip()
                                if not name or name in seen_programs:
                                    continue

                                version = "N/A"
                                try:
                                    version = winreg.QueryValueEx(subkey, "DisplayVersion")[0].strip()
                                except (WindowsError, KeyError):
                                    pass

                                # Skip system components and irrelevant entries
                                skip_keywords = [
                                    "update", "microsoft", "windows", "cache", "installer",
                                    "pack", "driver", "system", "component", "setup",
                                    "prerequisite", "runtime", "application", "sdk"
                                ]

                                if any(keyword in name.lower() for keyword in skip_keywords):
                                    continue

                                software_list.append({
                                    'name': name,
                                    'version': version
                                })
                                seen_programs.add(name)

                            except (WindowsError, KeyError):
                                continue

                        except WindowsError:
                            continue
                        finally:
                            if subkey is not None:
                                winreg.CloseKey(subkey)

                except WindowsError as e:
                    print(f"Error accessing {key_path}: {str(e)}")
                finally:
                    if reg_key is not None:
                        winreg.CloseKey(reg_key)

            # Sort the list by program name
            software_list.sort(key=lambda x: x['name'].lower())

            print(f"\nTotal programs found: {len(software_list)}")
            return {
                'status': 'success',
                'data': software_list
            }

        except Exception as e:
            error_msg = f"Error in software inventory: {str(e)}"
            print(f"ERROR: {error_msg}")
            return {
                'status': 'error',
                'message': error_msg
            }

    def handle_power_management(self, data):
        """Handle power management commands with enhanced functionality"""
        action = data.get('action')
        seconds = data.get('seconds')

        try:
            if platform.system() == 'Windows':
                if action == 'shutdown':
                    if seconds is not None:
                        if seconds > 0:
                            os.system(f'shutdown /s /t {seconds}')
                        else:
                            raise ValueError("Invalid shutdown time")
                    else:
                        os.system('shutdown /s /t 1')

                elif action == 'restart':
                    os.system('shutdown /r /t 1')

                elif action == 'lock':
                    os.system('rundll32.exe user32.dll,LockWorkStation')

                elif action == 'cancel_scheduled':
                    os.system('shutdown /a')

            else:
                # Linux/Unix commands
                if action == 'shutdown':
                    if seconds is not None:
                        os.system(f'shutdown -h +{seconds // 60}')  # Convert seconds to minutes for Linux
                    else:
                        os.system('shutdown -h now')
                elif action == 'restart':
                    os.system('shutdown -r now')
                elif action == 'lock':
                    os.system('loginctl lock-session')
                elif action == 'cancel_scheduled':
                    os.system('shutdown -c')

            return {
                'status': 'success',
                'message': f'Power management action {action} initiated successfully'
            }

        except Exception as e:
            logging.error(f"Power management error: {str(e)}")
            return {
                'status': 'error',
                'message': f'Failed to execute power action: {str(e)}'
            }

    def handle_command_execution(self, data):
        """Execute system commands"""
        command = data.get('command')
        try:
            result = subprocess.run(
                command,
                shell=True,
                capture_output=True,
                text=True,
                timeout=30
            )
            return {
                'status': 'success',
                'data': {
                    'stdout': result.stdout,
                    'stderr': result.stderr,
                    'return_code': result.returncode
                }
            }
        except Exception as e:
            return {'status': 'error', 'message': str(e)}

    def handle_network_monitor(self, data):
        """Monitor network statistics"""
        return {
            'status': 'success',
            'data': {
                'connections': [conn._asdict() for conn in psutil.net_connections()],
                'io_counters': dict(psutil.net_io_counters()._asdict())
            }
        }

    def handle_start_rdp(self, data):
        """Start RDP server with improved state management"""
        try:
            # Force close any existing RDP server
            if self.rdp_server is not None:
                try:
                    logging.info("Stopping existing RDP server before starting a new one")
                    self.rdp_server.stop()
                    self.rdp_server = None

                    # If there's a thread, wait for it to terminate
                    if self.rdp_thread and self.rdp_thread.is_alive():
                        self.rdp_thread.join(timeout=2)
                        self.rdp_thread = None

                    # Additional sleep to ensure resources are released
                    time.sleep(1)
                except Exception as e:
                    logging.error(f"Error stopping existing RDP server: {str(e)}")

            # Use the actual server's IP for the RDP connection
            rdp_host = socket.gethostbyname(socket.gethostname())
            rdp_port = 5900  # Default RDP port

            # Create and start RDP server
            self.rdp_server = RDPServer(host='0.0.0.0', port=rdp_port)
            self.rdp_thread = threading.Thread(target=self.rdp_server.start)
            self.rdp_thread.daemon = True
            self.rdp_thread.start()

            # Wait for server to start
            time.sleep(1.5)  # Increased wait time for server startup

            return {
                'status': 'success',
                'data': {
                    'ip': rdp_host,
                    'port': rdp_port
                }
            }

        except Exception as e:
            logging.error(f"Failed to start RDP server: {str(e)}")
            # Clean up in case of error
            if hasattr(self, 'rdp_server') and self.rdp_server:
                try:
                    self.rdp_server.stop()
                    self.rdp_server = None
                except:
                    pass
            return {
                'status': 'error',
                'message': str(e)
            }

    def handle_stop_rdp(self, data):
        """Stop RDP server with improved cleanup"""
        try:
            if self.rdp_server:
                logging.info("Stopping RDP server")
                self.rdp_server.stop()
                self.rdp_server = None

                # Clean up thread reference
                if self.rdp_thread:
                    if self.rdp_thread.is_alive():
                        self.rdp_thread.join(timeout=3)
                    self.rdp_thread = None

                return {'status': 'success', 'message': 'RDP server stopped successfully'}
            else:
                return {'status': 'success', 'message': 'No RDP server was running'}

        except Exception as e:
            logging.error(f"Error stopping RDP server: {str(e)}")
            # Force reset server state even if there's an error
            self.rdp_server = None
            self.rdp_thread = None
            return {'status': 'error', 'message': f'Error stopping RDP server: {str(e)}'}

    def stop(self):
        logging.info("Shutting down server...")
        self.running = False

        for client in list(self.clients.values()):
            try:
                client['socket'].close()
            except:
                pass

        try:
            self.server_socket.close()
        except:
            pass

        logging.info("Server shutdown complete")


if __name__ == "__main__":
    server = MCCServer()
    try:
        server.start()
    except KeyboardInterrupt:
        logging.info("Received shutdown signal")
        server.stop()
