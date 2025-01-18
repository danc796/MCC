import socket
import ssl
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


class MCCServer:
    def __init__(self, host='0.0.0.0', port=5000, cert_file='server.crt', key_file='server.key'):
        self.host = host
        self.port = port
        self.cert_file = cert_file
        self.key_file = key_file
        self.server_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        self.clients = {}

        # Create SSL context
        self.ssl_context = self.create_ssl_context()

        # Setup logging with enhanced security logging
        logging.basicConfig(
            filename='mcc_server.log',
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s'
        )

    def create_ssl_context(self):
        """Create and configure SSL context with modern security settings"""
        context = ssl.create_default_context(ssl.Purpose.CLIENT_AUTH)

        try:
            context.load_cert_chain(certfile=self.cert_file, keyfile=self.key_file)
        except (ssl.SSLError, FileNotFoundError) as e:
            logging.error(f"Failed to load SSL certificates: {str(e)}")
            raise

        context.minimum_version = ssl.TLSVersion.TLSv1_3
        context.maximum_version = ssl.TLSVersion.TLSv1_3
        context.set_ciphers('ECDHE+AESGCM:ECDHE+CHACHA20')

        # Make the server more permissive for development
        context.verify_mode = ssl.CERT_NONE

        return context

    def start(self):
        """Start the TLS-enabled server and listen for connections"""
        try:
            self.server_socket.bind((self.host, self.port))
            self.server_socket.listen(2)
            logging.info(f"Secure server started on {self.host}:{self.port}")

            while True:
                try:
                    client_socket, address = self.server_socket.accept()
                    logging.info(f"Received connection request from {address}")

                    # Wrap the socket with TLS
                    secure_socket = self.ssl_context.wrap_socket(
                        client_socket,
                        server_side=True,
                        do_handshake_on_connect=True
                    )
                    logging.info(f"TLS handshake completed with {address}")

                    # Generate encryption key
                    cipher_suite = Fernet(Fernet.generate_key())

                    # Start client handler thread
                    client_handler = threading.Thread(
                        target=self.handle_client,
                        args=(secure_socket, address, cipher_suite),
                        daemon=True
                    )
                    client_handler.start()
                    logging.info(f"Started handler thread for {address}")

                except ssl.SSLError as e:
                    logging.error(f"SSL handshake failed with {address}: {str(e)}")
                    client_socket.close()
                except Exception as e:
                    logging.error(f"Error accepting connection from {address}: {str(e)}")
                    if 'client_socket' in locals():
                        client_socket.close()

        except Exception as e:
            logging.error(f"Fatal server error: {str(e)}", exc_info=True)
            raise
        finally:
            self.server_socket.close()

    def handle_client(self, secure_socket, address, cipher_suite):
        """Handle individual client connections with TLS encryption"""
        try:
            # Send the encryption key
            encryption_key = Fernet.generate_key()
            secure_socket.sendall(encryption_key)

            # Create new cipher suite with the generated key
            cipher_suite = Fernet(encryption_key)

            self.clients[address] = {
                'socket': secure_socket,
                'cipher_suite': cipher_suite,
                'last_seen': datetime.now(),
                'system_info': self.get_system_info()
            }

            logging.info(f"Successfully established encrypted connection with {address}")

            while True:
                try:
                    encrypted_data = secure_socket.recv(4096)
                    if not encrypted_data:
                        break

                    data = cipher_suite.decrypt(encrypted_data).decode()
                    command = json.loads(data)
                    response = self.process_command(command)
                    encrypted_response = cipher_suite.encrypt(json.dumps(response).encode())
                    secure_socket.sendall(encrypted_response)

                except Exception as e:
                    logging.error(f"Communication error with {address}: {str(e)}")
                    break

        except Exception as e:
            logging.error(f"Connection handler error for {address}: {str(e)}")
        finally:
            self.clients.pop(address, None)
            secure_socket.close()

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
        """Process incoming commands from clients"""
        cmd_type = command.get('type')
        cmd_data = command.get('data', {})

        command_handlers = {
            'system_info': self.handle_system_info,
            'hardware_monitor': self.handle_hardware_monitor,
            'software_inventory': self.handle_software_inventory,
            'power_management': self.handle_power_management,
            'file_transfer': self.handle_file_transfer,
            'execute_command': self.handle_command_execution,
            'network_monitor': self.handle_network_monitor
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
        """Get installed software inventory without source tracking"""
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
                try:
                    reg_key = winreg.OpenKey(reg_root, key_path)

                    for i in range(winreg.QueryInfoKey(reg_key)[0]):
                        try:
                            subkey_name = winreg.EnumKey(reg_key, i)
                            with winreg.OpenKey(reg_key, subkey_name) as subkey:
                                try:
                                    # Try to get program name
                                    name = winreg.QueryValueEx(subkey, "DisplayName")[0].strip()
                                    if not name or name in seen_programs:
                                        continue

                                    # Try to get version
                                    try:
                                        version = winreg.QueryValueEx(subkey, "DisplayVersion")[0].strip()
                                    except:
                                        version = "N/A"

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
                except WindowsError as e:
                    print(f"Error accessing {key_path}: {str(e)}")
                finally:
                    try:
                        winreg.CloseKey(reg_key)
                    except:
                        pass

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

    def handle_file_transfer(self, data):
        """Handle file transfer operations"""
        operation = data.get('operation')
        filepath = data.get('filepath')

        if operation == 'receive':
            try:
                with open(filepath, 'rb') as file:
                    return {
                        'status': 'success',
                        'data': file.read().decode('utf-8')
                    }
            except Exception as e:
                return {'status': 'error', 'message': str(e)}

        return {'status': 'error', 'message': 'Invalid file operation'}

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


if __name__ == "__main__":
    server = MCCServer()
    server.start()
