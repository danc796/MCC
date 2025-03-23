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
from serverRDP import RDPServer
import time


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

            # Get scale parameter from client (if provided)
            scale = data.get('scale', 1.0)

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
                    'port': rdp_port,
                    'scale': scale
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
