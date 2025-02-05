import platform
import psutil
import socket
import cpuinfo
import datetime
import GPUtil
from tabulate import tabulate
import sys
import distro
import warnings

warnings.filterwarnings('ignore')


class SystemInformation:
    def __init__(self):
        self.data = {}
        self.current_time = datetime.datetime.now()

    def get_datetime_info(self):
        """Get current date and time information"""
        self.data['datetime'] = {
            'Date': self.current_time.strftime('%Y-%m-%d'),
            'Time': self.current_time.strftime('%H:%M:%S'),
            'Timezone': datetime.datetime.now().astimezone().tzname(),
            'Unix Timestamp': int(self.current_time.timestamp())
        }

    def get_system_info(self):
        """Gather basic system information"""
        self.data['system'] = {
            'OS': platform.system(),
            'OS Version': platform.version(),
            'OS Release': platform.release(),
            'Machine': platform.machine(),
            'Processor': platform.processor(),
            'Architecture': platform.architecture(),
            'Python Version': sys.version,
            'Hostname': socket.gethostname(),
        }

        # Add Linux-specific information if available
        if platform.system() == 'Linux':
            self.data['system'].update({
                'Distro': f"{distro.name()} {distro.version()}",
                'Distro ID': distro.id(),
                'Distro Codename': distro.codename()
            })

    def get_cpu_info(self):
        """Gather detailed CPU information"""
        cpu = cpuinfo.get_cpu_info()
        self.data['cpu'] = {
            'Brand': cpu['brand_raw'],
            'Architecture': cpu['arch'],
            'Bits': cpu['bits'],
            'Cores (Physical)': psutil.cpu_count(logical=False),
            'Cores (Logical)': psutil.cpu_count(logical=True),
            'Current Frequency': f"{psutil.cpu_freq().current:.2f} MHz",
            'CPU Usage Per Core': [f"{x}%" for x in psutil.cpu_percent(percpu=True)],
            'Total CPU Usage': f"{psutil.cpu_percent()}%"
        }

    def get_memory_info(self):
        """Gather RAM and swap information"""
        memory = psutil.virtual_memory()
        swap = psutil.swap_memory()

        def bytes_to_gb(bytes_val):
            return f"{bytes_val / (1024 ** 3):.2f} GB"

        self.data['memory'] = {
            'Total RAM': bytes_to_gb(memory.total),
            'Available RAM': bytes_to_gb(memory.available),
            'Used RAM': bytes_to_gb(memory.used),
            'RAM Usage': f"{memory.percent}%",
            'Total Swap': bytes_to_gb(swap.total),
            'Used Swap': bytes_to_gb(swap.used),
            'Free Swap': bytes_to_gb(swap.free),
            'Swap Usage': f"{swap.percent}%"
        }

    def get_disk_info(self):
        """Gather storage information for all partitions"""
        self.data['disks'] = []

        for partition in psutil.disk_partitions():
            try:
                usage = psutil.disk_usage(partition.mountpoint)
                self.data['disks'].append({
                    'Device': partition.device,
                    'Mountpoint': partition.mountpoint,
                    'File System': partition.fstype,
                    'Total Size': f"{usage.total / (1024 ** 3):.2f} GB",
                    'Used': f"{usage.used / (1024 ** 3):.2f} GB",
                    'Free': f"{usage.free / (1024 ** 3):.2f} GB",
                    'Usage': f"{usage.percent}%"
                })
            except PermissionError:
                continue

    def get_gpu_info(self):
        """Gather GPU information using GPUtil"""
        try:
            gpus = GPUtil.getGPUs()
            self.data['gpu'] = []

            for gpu in gpus:
                self.data['gpu'].append({
                    'ID': gpu.id,
                    'Name': gpu.name,
                    'Load': f"{gpu.load * 100}%",
                    'Free Memory': f"{gpu.memoryFree} MB",
                    'Used Memory': f"{gpu.memoryUsed} MB",
                    'Total Memory': f"{gpu.memoryTotal} MB",
                    'Temperature': f"{gpu.temperature} °C",
                    'UUID': gpu.uuid
                })
        except Exception as e:
            self.data['gpu'] = f"GPU information unavailable: {str(e)}"

    def get_battery_info(self):
        """Gather battery information if available"""
        try:
            battery = psutil.sensors_battery()
            if battery:
                self.data['battery'] = {
                    'Percentage': f"{battery.percent}%",
                    'Power Plugged': battery.power_plugged,
                    'Time Left': str(datetime.timedelta(seconds=battery.secsleft)) if battery.secsleft != -1 else "N/A"
                }
            else:
                self.data['battery'] = "No battery detected"
        except Exception:
            self.data['battery'] = "Battery information unavailable"

    def collect_all_info(self):
        """Collect all system information"""
        self.get_datetime_info()
        self.get_system_info()
        self.get_cpu_info()
        self.get_memory_info()
        self.get_disk_info()
        self.get_gpu_info()
        self.get_battery_info()

    def display_info(self):
        """Display the collected information in a formatted manner"""
        output = []
        for section, data in self.data.items():
            output.append(f"\n{'=' * 20} {section.upper()} {'=' * 20}\n")

            if isinstance(data, str):
                # Handle string data (like error messages)
                output.append(data)
            elif isinstance(data, list):
                # Handle lists of dictionaries (like disks and GPU)
                if data and isinstance(data[0], dict):
                    headers = list(data[0].keys()) if data else []
                    output.append(tabulate([item.values() for item in data],
                                           headers=headers,
                                           tablefmt='grid'))
            elif isinstance(data, dict):
                # Handle nested dictionaries
                if any(isinstance(v, dict) for v in data.values()):
                    for subsection, subdata in data.items():
                        output.append(f"\n{'-' * 10} {subsection} {'-' * 10}")
                        if isinstance(subdata, list):
                            for item in subdata:
                                if isinstance(item, dict):
                                    headers = list(item.keys())
                                    output.append(tabulate([item.values()],
                                                           headers=headers,
                                                           tablefmt='grid'))
                        elif isinstance(subdata, dict):
                            output.append(tabulate([[k, v] for k, v in subdata.items()],
                                                   headers=['Property', 'Value'],
                                                   tablefmt='grid'))
                else:
                    output.append(tabulate([[k, v] for k, v in data.items()],
                                           headers=['Property', 'Value'],
                                           tablefmt='grid'))
            else:
                output.append(str(data))

        # Join all output parts into a single string
        return '\n'.join(output)


def main():
    """Main function to run the system information collection"""
    try:
        sys_info = SystemInformation()
        sys_info.collect_all_info()

        # Generate output
        output = sys_info.display_info()

        # Write to file
        current_time = datetime.datetime.now()
        pc_name = socket.gethostname()
        filename = f"{pc_name}_info_{current_time.strftime('%d.%m_%H;%M')}.txt"
        # Replace forward slashes with underscores for valid filename
        filename = filename.replace('/', '_')
        with open(filename, 'w', encoding='utf-8') as f:
            f.write(output)

        print(f"\nSystem information has been saved to {filename}")

    except Exception as e:
        print(f"An error occurred: {str(e)}")


if __name__ == "__main__":
    main()
