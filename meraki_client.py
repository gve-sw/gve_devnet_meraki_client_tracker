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

import meraki
from meraki import APIError
from rich.console import Console

from config import *

# Rich Console Instance
console = Console()

# Meraki Dashboard Instance
dashboard = meraki.DashboardAPI(api_key=MERAKI_API_KEY, suppress_logging=True)


def get_network_ids(org_name):
    """
    Get network IDs in org
    :param org_name: Org Name
    :return: Network ID
    """
    # Get Meraki Org ID
    orgs = dashboard.organizations.getOrganizations()
    org_id = ''
    for org in orgs:
        if org['name'] == org_name:
            org_id = org['id']
            break

    if org_id == '':
        return None

    # Get Meraki Network IDs
    networks = dashboard.organizations.getOrganizationNetworks(organizationId=org_id)
    net_ids = [(net_id['id'], net_id['name']) for net_id in networks]

    return net_ids


def sorted_list_network_names(network_ids):
    """
    Create list of sorted network names from usage dictionary
    :param network_ids: tuple list of networks and ids
    :return: List of sorted Network Names
    """
    # Create list of network names
    network_names = [network[1] for network in network_ids]

    # Sort networks alphabetically
    network_names = sorted(network_names, key=lambda d: d.lower())

    return network_names


class MerakiClientInfo:
    def __init__(self, mac, time_period):
        self.mac = mac
        self.time_period = time_period
        self.net_ids = get_network_ids(ORG_NAME)
        self.sorted_net_names = sorted_list_network_names(self.net_ids)
        self.clientDetails = None
        self.usage = None

    def client_detail_history(self):
        # Set Client Details for client across networks (network specific)

        # Build client data dictionary for each network and summary
        client_details = {"client_mac": self.mac, "networks": []}

        # if no mac is found, return empty details dictionary
        if not self.mac:
            console.print(f'[red]MAC Address is empty... skipping.')
            self.clientDetails = client_details
            return

        for network in self.net_ids:
            try:
                # Get client details in network for specific mac at specific time range
                response = dashboard.networks.getNetworkClients(
                    network[0], mac=self.mac, timespan=self.time_period, total_pages='all'
                )
            except APIError as a:
                if 'not found' in a.message['errors'][0]:
                    console.print(f'[red]Client Not Found [/] in {network[1]}.')

                    # Log Network, and empty applications list
                    client_details['networks'].append({"network_name": network[1], "details": {}})
                    continue
                else:
                    return

            if len(response) > 0:
                console.print(f"Found Client Details Data in [blue]{network[1]}![/]")

                # build new client details dictionary containing only relevant info
                client_details_minimized = {
                    "description": response[0]['description'],
                    "ip": response[0]['ip'],
                    "mac": response[0]['mac'],
                    "user": response[0]['user'],
                    'manufacturer': response[0]['manufacturer'],
                    'os': response[0]['os'],
                    'recentDeviceSerial': response[0]['recentDeviceSerial'],
                    'recentDeviceName': response[0]['recentDeviceName'],
                    'recentDeviceConnection': response[0]['recentDeviceConnection'],
                    'status': response[0]['status'],
                    'vlan': response[0]['vlan']
                }

                # if wireless, include ssid, else include switch port
                if client_details_minimized['recentDeviceConnection'] == 'Wireless':
                    client_details_minimized['ssid'] = response[0]['ssid']
                else:
                    client_details_minimized['switchport'] = response[0]['switchport']

                net_client_details = {"network_name": network[1], "client_details": client_details_minimized}
                client_details['networks'].append(net_client_details)

        self.clientDetails = client_details

    def app_usage_history(self):
        # Return App usage history for client across networks (summary and network specific)

        # Build app usage dictionary for each network and summary
        app_usage = {"client_mac": self.mac, "summary": {}, "networks": []}

        # if no mac is found, return empty details dictionary
        if not self.mac:
            console.print(f'[red]MAC Address is empty... skipping.')
            self.usage = app_usage
            return

        for network in self.net_ids:
            try:
                response = dashboard.networks.getNetworkClientsApplicationUsage(
                    network[0], self.mac, timespan=self.time_period, total_pages='all'
                )
            except APIError as a:
                if 'not found' in a.message['errors'][0]:
                    console.print(f'[red]Client Not Found [/] in {network[1]}.')

                    # Log Network, and empty applications list
                    app_usage['networks'].append({"network_name": network[1], "applications": {}})
                    continue
                else:
                    return

            # Sort returned applications alphabetically
            applications = sorted(response[0]['applicationUsage'], key=lambda d: d['application'].lower())

            console.print(
                f"Found usage Data in [blue]{network[1]}[/] for [yellow]{len(applications)} applications![/]")

            # Summarize usage data across networks, track data per network
            net_app_usage = {"network_name": network[1], "applications": {}}

            for application in applications:
                name = application['application']

                # Append to Network Dictionary
                net_app_usage['applications'][name] = [application['received'], application['sent']]

                # Append to Summary Dictionary
                if name not in app_usage['summary']:
                    app_usage['summary'][name] = [application['received'], application['sent']]
                else:
                    app_usage['summary'][name][0] += application['received']
                    app_usage['summary'][name][1] += application['sent']

            app_usage['networks'].append(net_app_usage)

        self.usage = app_usage
