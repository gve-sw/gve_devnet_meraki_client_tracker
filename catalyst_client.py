#!/usr/bin/env python3
"""
Copyright (c) 2023 Cisco and/or its affiliates.
This software is licensed to you under the terms of the Cisco Sample
Code License, Version 1.1 (the "License"). You may obtain a copy of the
License at
https://developer.cisco.com/docs/licenses
All use of the material herein must be in accordance with the terms of
the License. All rights not expressly granted by the License are
reserved. Unless required by applicable law or agreed to separately in
writing, software distributed under the License is distributed on an "AS
IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express
or implied.
"""

__author__ = "Trevor Maco <tmaco@cisco.com>"
__copyright__ = "Copyright (c) 2023 Cisco and/or its affiliates."
__license__ = "Cisco Sample Code License, Version 1.1"

from netmiko import ConnectHandler
from rich.console import Console

# Rich Console Instance
console = Console()


def execute_switch_commands(connection, commands):
    """
    Connect to target switch over ssh, execute command
    :param connection: Netmiko ssh connection object
    :param commands: CLI command to execute
    :return: String containing results of 'show run' command
    """
    # Enter privilege mode
    connection.enable()

    # Send command
    output = connection.send_command(commands, use_textfsm=True)

    # Check if show output is valid
    if 'Invalid' in output:
        console.print(f' - [red]Failed to execute "{commands}"[/], please ensure the command(s) are correct!')
        return None
    else:
        console.print(f' - Executed [blue]"{commands}"[/] successfully!')
        return output


def convert_mac(mac):
    """
    Convert mac address to mac the Catalyst switch understands
    :param mac: mac address in meraki format
    :return: mac address in catalyst format
    """
    converted_address = mac.replace(':', '').lower()
    converted_address = ".".join([converted_address[i:i + 4] for i in range(0, len(converted_address), 4)])
    return converted_address


class CatalystClientInfo:
    def __init__(self, mac, ip):
        self.switch_hostname = None
        self.vlan = None
        self.interface = None
        self.interface_status = {}
        self.cdp = []
        self.lldp = []
        self.mac = convert_mac(mac) if mac else None
        self.lan_ip = ip
        self.switch_connection = None

    def connectToSwitch(self, device_info):
        # Start SSH session and login to device
        console.print(f'Connecting to Switch at [green]{device_info["ip"]}[/]...')
        try:
            self.switch_connection = ConnectHandler(**device_info)
        except Exception:
            console.print(f'Unable to connect to switch [red]{device_info["ip"]}[/], skipping...')
            return None

        console.print(f'[green]Connected to switch {device_info["ip"]}![/]')
        return self.switch_connection

    def disconnectFromSwitch(self):
        # End SSH session
        self.switch_connection.disconnect()

    def clientPresentCheck(self):
        if self.mac:
            # Check if Mac is present in Mac Address Table (indicates this is the right switch)
            output = execute_switch_commands(self.switch_connection, f'show mac address-table | include {self.mac}')
        else:
            # Check if ip is present in ARP Table (indicates this is the right switch)
            output = execute_switch_commands(self.switch_connection, f'show ip arp | include {self.lan_ip}')

        return output != ''

    def hostname(self):
        # Get hostname of switch
        output = execute_switch_commands(self.switch_connection, 'show run | i hostname')

        if output:
            self.switch_hostname = output.strip().split()[1]

    def macAddressTable(self):
        # Get mac address table information
        output = execute_switch_commands(self.switch_connection, f'show mac address-table | include {self.mac}')
        output = output.strip().split()

        # Set new parameters based on Mac address table
        self.vlan = output[0]
        self.interface = output[3]

    def arpTable(self):
        if self.mac:
            # Get show arp table information (for ip)
            output = execute_switch_commands(self.switch_connection, f'show ip arp | i {self.mac}')

            if output and len(output) > 0:
                self.lan_ip = output[0]['address']
        else:
            # Get show arp table information (for mac)
            output = execute_switch_commands(self.switch_connection, f'show ip arp | i {self.lan_ip}')

            if output and len(output) > 0:
                self.mac = output[0]['mac']

    def interfaceStatus(self):
        # Get show interface status information, set status
        output = execute_switch_commands(self.switch_connection, f'show ip int br {self.interface}')

        if output and len(output) > 0:
            self.interface_status['status'] = output[0]['status'] + '/' + output[0]['proto']

        # Get show interface status information, set status
        output = execute_switch_commands(self.switch_connection, f'show int {self.interface} trunk')

        if output:
            # Split and remove empty entries
            output = output.split('\n')
            output = [x for x in output if x]

            for line in output:
                if not line.startswith('Port'):
                    line = line.split()

                    if len(line) == 5:
                        # Interface not trunking
                        if line[1] == 'off':
                            self.interface_status['mode'] = 'access'
                        else:
                            self.interface_status['mode'] = 'trunk'

                        self.interface_status['encapsulation'] = line[2]
                        self.interface_status['native_vlan'] = line[4]

                    elif len(line) == 2 and "allowed_vlans" not in self.interface_status:
                        self.interface_status['allowed_vlans'] = line[1]

    def neighborInformation(self):
        # Get show interface status information, set status
        output = execute_switch_commands(self.switch_connection, f'show cdp neighbors')

        if output and len(output) > 0:
            self.cdp = output

        # Get show interface status information, set status
        output = execute_switch_commands(self.switch_connection, f'show lldp neighbors')

        if output and len(output) > 0:
            self.lldp = output
