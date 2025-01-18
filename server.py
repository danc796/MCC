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


class MCCServer:
    def __init__(self, host='0.0.0.0', port=5000):
        self.host = host
        self.port = port
        self.server_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        self.clients = {}
        self.encryption_key = Fernet.generate_key()
        self.cipher_suite = Fernet(self.encryption_key)

        # Setup logging
        logging.basicConfig(
            filename='mcc_server.log',
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s'
        )

    def start(self):
        """Start the server and listen for connections"""
        try:
            self.server_socket.bind((self.host, self.port))
            self.server_socket.listen(2)
            logging.info(f"Server started on {self.host}:{self.port}")

            while True:
                client_socket, address = self.server_socket.accept()
                client_handler = threading.Thread(
                    target=self.handle_client,
                    args=(client_socket, address)
                )
                client_handler.start()

        except Exception as e:
            logging.error(f"Server error: {str(e)}")
            raise

    def handle_client(self, client_socket, address):
        """Handle individual client connections"""
        try:
            # Send encryption key immediately after connection
            client_socket.send(self.encryption_key)

            self.clients[address] = {
                'socket': client_socket,
                'last_seen': datetime.now(),
                'system_info': self.get_system_info()
            }
            logging.info(f"New connection from {address}, encryption key sent")

            while True:
                encrypted_data = client_socket.recv(4096)
                if not encrypted_data:
                    break

                data = self.cipher_suite.decrypt(encrypted_data).decode()
                command = json.loads(data)
                response = self.process_command(command)

                encrypted_response = self.cipher_suite.encrypt(
                    json.dumps(response).encode()
                )
                client_socket.send(encrypted_response)

        except Exception as e:
            logging.error(f"Error handling client {address}: {str(e)}")
        finally:
            self.clients.pop(address, None)
            client_socket.close()

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
        schedule_time = data.get('schedule_time')

        try:
            if platform.system() == 'Windows':
                if action == 'shutdown':
                    if schedule_time:
                        # Convert schedule_time to seconds
                        seconds = int((datetime.strptime(schedule_time, '%Y-%m-%d %H:%M:%S') -
                                       datetime.now()).total_seconds())
                        if seconds > 0:
                            os.system(f'shutdown /s /t {seconds}')
                        else:
                            raise ValueError("Scheduled time must be in the future")
                    else:
                        os.system('shutdown /s /t 1')

                elif action == 'restart':
                    os.system('shutdown /r /t 1')

                elif action == 'lock':
                    os.system('rundll32.exe user32.dll,LockWorkStation')

                elif action == 'cancel_scheduled':
                    os.system('shutdown /a')

            else:
                if action == 'shutdown':
                    if schedule_time:
                        os.system(f'shutdown -h {schedule_time}')
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
