import wmi
import get_qlik_sense
import argparse
import socket

parser = argparse.ArgumentParser()
parser.add_argument('--server', help='Qlik Sense Server to connect to.')
parser.add_argument('--certs', help='Path to certificates.')
parser.add_argument('--user', help='Username in format domain\\user')
parser.add_argument('--wmi', help='True/False, set to True to collect WMI information from servers This switch requires windows authentication (username and password)')
parser.add_argument('--password', help='Password of user.')
args = parser.parse_args()

lhost = socket.gethostname().upper()
lfqdn = socket.getfqdn().upper()

class Connect:

    def __init__(self, user, password ):
        self.user = user
        self.password = password

    def get_os(self, server):
        if server.upper() == lhost or server.upper() == lfqdn:
            c = wmi.WMI()
            for os in c.Win32_OperatingSystem():
                return os.Caption
        else:
            c = wmi.WMI(server, user=self.user, password=self.password)
            for os in c.Win32_OperatingSystem():
                return os.Caption

    def get_memory(self, server):
        if server.upper() == lhost or server.upper() == lfqdn:
            c = wmi.WMI()
            for memory in c.Win32_OperatingSystem():
                return str(format(int(memory.FreePhysicalMemory)/1000000, '.2f')+'Gb')
        else:
            c = wmi.WMI(server, user=self.user, password=self.password)
            for memory in c.Win32_OperatingSystem():
                return str(format(int(memory.FreePhysicalMemory)/1000000, '.2f')+'Gb')

    def get_totalram(self, server):
        if server.upper() == lhost or server.upper() == lfqdn:
            c = wmi.WMI()
            ram = 0
            for memory in c.Win32_PhyscialMemory():
                ram += int(memory.Capacity)
        else:
            c = wmi.WMI(server, user=self.user, password=self.password)
            ram = 0
            for memory in c.Win32_PhysicalMemory():
                ram += int(memory.Capacity)
        return str(format(ram / 1024 ** 3, '.2f')+'Gb')


    def get_core(self, server):
        dcpu = {}
        if server.upper() == lhost or server.upper() == lfqdn:
            c = wmi.WMI()
            cores = 0
            for cpu in c.Win32_Processor():
                cores += 1
                dcpu['number of cpu'] = cores     
        else:
            c = wmi.WMI(server, user=self.user, password=self.password)
            cores = 0
            for cpu in c.Win32_Processor():
                cores += 1
                dcpu['number of cpu'] = cores
        return dcpu

    def get_cpuname(self,server):
        dcpu = {}
        if server.upper() == lhost or server.upper() == lfqdn:
            c = wmi.WMI()
            for cpu in c.Win32_Processor():
                dcpu['cpu name'] = cpu.Name
        else:
            c = wmi.WMI(server, user=self.user, password=self.password)
            for cpu in c.Win32_Processor():
                dcpu['cpu name'] = cpu.Name
        return dcpu

    def get_logicalcore(self, server):
        dcpu = {}
        if server.upper() == lhost or server.upper() == lfqdn:
            c = wmi.WMI()
            cores = 0
            for cpu in c.Win32_Processor():
                cores += cpu.NumberofLogicalProcessors
                dcpu['number of logical cores'] = cores
        else:
            c = wmi.WMI(server, user=self.user, password=self.password)
            cores = 0
            for cpu in c.Win32_Processor():
                cores += cpu.NumberofLogicalProcessors
                dcpu['number of logical cores'] = cores

        return dcpu
    
    def get_disk(self, server):
        disks = {}
        if server.upper() == lhost or server.upper() == lfqdn:
            c = wmi.WMI()
            for disk in c.Win32_LogicalDisk (DriveType=3):
                disks[disk.DeviceID] = str(format(int(disk.FreeSpace)/1000000000,'.2f')+'Gb')
        else:
            c = wmi.WMI(server, user=self.user, password=self.password)
            for disk in c.Win32_LogicalDisk (DriveType=3):
                disks[disk.DeviceID] = str(format(int(disk.FreeSpace)/1000000000,'.2f')+'Gb')
        return disks

    def get_totaldisk(self,server):
        disks = {}
        if server.upper() == lhost or server.upper() == lfqdn:
            c = wmi.WMI()
            for disk in c.Win32_LogicalDisk (DriveType=3):
                disks[disk.DeviceID] = str(format(int(disk.Size)/1000000000,'.2f')+'Gb')
        else:
            c = wmi.WMI(server, user=self.user, password=self.password)
            for disk in c.Win32_LogicalDisk (DriveType=3):
                disks[disk.DeviceID] = str(format(int(disk.Size)/1000000000,'.2f')+'Gb')
        return (disks)


    def wmi_results(self):
        server_details = get_qlik_sense.get_servernode()
        site = []
        for server in range(len(server_details)):
            site.append (server_details[server][1])
        node_details = {}
        for item in range(len(site)):
            node_details[site[item]] = {'os':auth.get_os(site[item])}
            node_details[site[item]].update({'ram': auth.get_memory(site[item])})
            node_details[site[item]].update({'total ram': auth.get_totalram(site[item])})
            node_details[site[item]].update({'disk free': auth.get_disk(site[item])})
            node_details[site[item]].update({'disk total': auth.get_totaldisk(site[item])})
            node_details[site[item]].update(auth.get_core(site[item]))
            node_details[site[item]].update(auth.get_cpuname(site[item]))
            node_details[site[item]].update(auth.get_logicalcore(site[item]))

        servers = []
        for key in node_details.items():
            servers.append(key[0])


        wmi_data = []
        for item in range(len(servers)):
            wmi_data.append ([node_details[servers[item]]['os'],
                node_details[servers[item]]['ram'],
                node_details[servers[item]]['total ram'],
                node_details[servers[item]]['disk free'],
                node_details[servers[item]]['disk total'],
                node_details[servers[item]]['cpu name'],
                node_details[servers[item]]['number of cpu'],
                node_details[servers[item]]['number of logical cores'],
                servers[item]]
                )
        return wmi_data

auth = Connect(user=args.user, password=args.password)


