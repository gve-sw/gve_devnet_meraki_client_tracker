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

import datetime
import threading
from io import BytesIO

import requests
import xlsxwriter
from flask import Flask, request, render_template, Response, jsonify
from rich.console import Console
from rich.panel import Panel

from catalyst_client import CatalystClientInfo
import config
from meraki_client import MerakiClientInfo

# Flask Config
app = Flask(__name__)
app.config['JSON_SORT_KEYS'] = False

# Form Submission Results
meraki_details = None
cat_details = None

# Rich Console Instance
console = Console()

# Global lock
lock = threading.Lock()

# Global progress variable
progress = 0


def catalyst_client_information(mac, ip, switch):
    """
    Build Catalyst Client Object, contains details
    :param mac: Client MAC
    :param ip: Client IP Address
    :param switch: Current Switch Connection Info (netmiko)
    :return:
    """
    global cat_details

    # Determine if switch contains target, extract information
    cat_client = CatalystClientInfo(mac, ip)

    # Connect to device via ssh
    connection = cat_client.connectToSwitch(switch)

    if connection:
        # Check if client present on switch
        present = cat_client.clientPresentCheck()

        # Client found
        if present:
            # Extract information about client
            cat_client.hostname()
            cat_client.arpTable()
            cat_client.macAddressTable()
            cat_client.interfaceStatus()
            cat_client.neighborInformation()

            # Disconnect from client
            cat_client.disconnectFromSwitch()

            # Set client equal to global variable (once the thread has the lock)
            with lock:
                cat_details = cat_client


def meraki_client_information(mac, time_period):
    """
    Build Meraki Client Object, contains details and usage data
    :param mac: Client Mac Address
    :param time_period: Time Period for data query
    :return:
    """
    global meraki_details, progress

    # Get Meraki client information and usage
    meraki_client = MerakiClientInfo(mac, time_period)

    console.print(Panel.fit(f"Getting Meraki Client Details", title="Step 2"))

    # Get meraki client details for client mac address across all networks
    meraki_client.client_detail_history()

    progress = 50

    console.print(Panel.fit(f"Getting App Usage History", title="Step 3"))

    # Get application information for client mac address across all networks
    meraki_client.app_usage_history()

    progress = 75

    meraki_details = meraki_client


# Methods
def getSystemTimeAndLocation():
    """Returns location and time of accessing device"""
    # request user ip
    userIPRequest = requests.get('https://get.geojs.io/v1/ip.json')
    userIP = userIPRequest.json()['ip']

    # request geo information based on ip
    geoRequestURL = 'https://get.geojs.io/v1/ip/geo/' + userIP + '.json'
    geoRequest = requests.get(geoRequestURL)
    geoData = geoRequest.json()

    # create info string
    location = geoData['country']
    timezone = geoData['timezone']
    current_time = datetime.datetime.now().strftime("%d %b %Y, %I:%M %p")
    timeAndLocation = "System Information: {}, {} (Timezone: {})".format(location, current_time, timezone)

    return timeAndLocation


def convert_to_sec(time_period):
    """
    Convert time period from submission form to seconds (required by Meraki API usage call)
    :param time_period: Specified in submission form (24h, 72h, 1 week, or custom)
    :return: Time period in seconds
    """
    # Default Case (return 24 hours)
    if time_period == '':
        return 24 * 3600
    # Hours case
    elif 'Hours' in time_period:
        hour = int(time_period.split(' ')[0])
        return hour * 3600
    elif 'Week' in time_period:
        week = int(time_period.split(' ')[0])
        return week * 7 * 24 * 3600
    # Custom interval case (hours)
    else:
        return int(time_period) * 3600


# Flask Routes
@app.route('/')
def index():
    """
    Homepage: Clear global variables on load
    :return:
    """
    global cat_details, meraki_details, progress

    # Clear global objects
    cat_details = None
    meraki_details = None

    progress = 0

    return render_template('index.html', hiddenLinks=False, timeAndLocation=getSystemTimeAndLocation(),
                           table_flag=False)


@app.route('/display', methods=["POST"])
def submit():
    """
    Handle form submission: query Meraki for application usage data for each network, generate summary table,
    return table information to webpage
    :return: A List of tables containing usage data (paginated at 1 by default)
    """
    global cat_details, meraki_details, progress

    # Clear global objects
    cat_details = None
    meraki_details = None

    progress = 0

    mac_address = request.form['mac_address']
    ip_address = request.form['ip_address']
    time_period = request.form['time_period']
    custom_period = request.form['custom-interval']

    console.print(Panel.fit("Submission Detected:"))
    console.print(
        f"For mac: [blue]{mac_address}[/], or optional ip: [yellow]{ip_address}[/], with Time Period: [yellow]{time_period}[/], and optional Custom "
        f"Period: [yellow]{custom_period}[/]")

    # Select custom value if present
    if len(custom_period) != 0:
        seconds = convert_to_sec(custom_period)
    else:
        seconds = convert_to_sec(time_period)

    # Set MAC address to none if IP provided, else set IP to None
    if len(ip_address) != 0:
        mac_address = None
    else:
        ip_address = None

    console.print(Panel.fit(f"Getting Catalyst Client Details", title="Step 1"))

    # Get catalyst client details for client mac address (disconnect from switch if data is found) - search all switches
    threads = []
    for switch in config.SWITCH_INFO:
        t = threading.Thread(target=catalyst_client_information, args=(mac_address, ip_address, switch,))
        threads.append(t)

    # Start all threads
    for x in threads:
        x.start()

    # Wait for all of them to finish
    for x in threads:
        x.join()

    progress = 25

    # Extract local mac and convert if only IP is provided
    if not mac_address and cat_details:
        mac_address = cat_details.mac.replace('.', '')
        mac_address = ":".join([mac_address[i:i + 2] for i in range(0, len(mac_address), 2)])

    # Get Meraki Client Information
    meraki_client_information(mac_address, seconds)

    console.print(Panel.fit(f"Constructing Usage and Client Data Tables", title="Step 4"))

    # Catalyst Details Section
    if cat_details:
        # CDP Table
        cat_details_cdp = ('cdp', cat_details.cdp)

        # LLDP Table
        cat_details_lldp = ('lldp', cat_details.lldp)

    else:
        cat_details_cdp = None
        cat_details_lldp = None

    # Meraki Details Section
    # Get details sorted correctly
    network_client_details = []
    for network in meraki_details.sorted_net_names:
        target_data = {}
        for net in meraki_details.clientDetails['networks']:
            if net['network_name'] == network:
                target_data = net
                break

        if len(target_data) > 0:
            network_client_details.append((network, target_data['client_details']))
        else:
            network_client_details.append((network, None))

    # Usage Section
    summary_applications = ('summary', meraki_details.usage['summary'], meraki_details.usage_pie_chart['summary'])

    # Add all other network information
    network_applications = []
    for network in meraki_details.sorted_net_names:
        for net in meraki_details.usage['networks']:
            if net['network_name'] == network:
                target_usage = net['applications']
                break

        for net in meraki_details.usage_pie_chart['networks']:
            if net['network_name'] == network:
                target_usage_pie_chart = net['applications']
                break

        network_applications.append((network, target_usage, target_usage_pie_chart))

    progress = 100

    # Render template with pagination links and data for the requested page
    return render_template('index.html', hiddenLinks=False, timeAndLocation=getSystemTimeAndLocation(), table_flag=True,
                           network_names=meraki_details.sorted_net_names,
                           mac_address=meraki_details.mac,
                           summary_table=summary_applications, network_tables=network_applications,
                           details_tables=network_client_details, cat_details_table=cat_details,
                           cat_details_table_cdp=cat_details_cdp, cat_details_table_lldp=cat_details_lldp)


@app.route('/progress')
def get_progress():
    """
    Get current process progress for progress bar display
    :return:
    """
    global progress

    # Return the progress as a JSON response
    return jsonify({'progress': progress})


@app.route('/download/client/catalyst')
def download_catalyst_client():
    """
    Download Excel file containing catalyst client information seen in WebGUI Tables (event triggered by download button)
    :return: Excel File
    """
    console.print(f"Downloading Excel File [green]catalyst_client_details.xlsx[/]")

    # Create in-memory file for writing Excel data
    output = BytesIO()

    # Create Excel workbook and add sheets
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    sheets = []

    # Catalyst Client Sheets
    sheet = workbook.add_worksheet('Catalyst Client Details')
    sheets.append(sheet)

    # Define Column Headers (Catalyst Client Details)
    fields = ['Field', 'Value']
    header_format = workbook.add_format({'bold': True, 'bottom': 2})

    # Write column headers
    sheet.write_row(0, 0, fields, header_format)

    # Write Column Rows
    if cat_details:
        sheet.write_row(1, 0, ['Status', 'Online'])
        sheet.write_row(2, 0, ['Mac Address', cat_details.mac if cat_details.mac else 'None'])
        sheet.write_row(3, 0, ['IP Address', cat_details.lan_ip if cat_details.lan_ip else 'None'])
        sheet.write_row(4, 0, ['VLAN', cat_details.vlan if cat_details.vlan else 'None'])
        sheet.write_row(5, 0,
                        ['Switch Hostname', cat_details.switch_hostname if cat_details.switch_hostname else 'None'])
        sheet.write_row(6, 0,
                        ['Interface', cat_details.interface if cat_details.interface else 'None'])
        sheet.write_row(7, 0,
                        ['Interface Status',
                         cat_details.interface_status['status'] if cat_details.interface_status['status'] else 'None'])
        sheet.write_row(8, 0,
                        ['Interface Mode',
                         cat_details.interface_status['mode'] if cat_details.interface_status['mode'] else 'None'])
        sheet.write_row(9, 0,
                        ['Allowed VLANs', cat_details.interface_status['allowed_vlans'] if cat_details.interface_status[
                            'allowed_vlans'] else 'None'])

        sheet = workbook.add_worksheet('CDP Table')
        sheets.append(sheet)

        # Define Column Headers (Catalyst Client Details)
        fields = ['Neighbor', 'Local Interface', 'Capability', 'Platform', 'Neighbor Interface']
        header_format = workbook.add_format({'bold': True, 'bottom': 2})

        # Write column headers
        sheet.write_row(0, 0, fields, header_format)

        for j, val in enumerate(cat_details.cdp):
            sheet.write_row(j + 1, 0, list(val.values()))

        sheet = workbook.add_worksheet('LLDP Table')
        sheets.append(sheet)

        fields = ['Neighbor', 'Local Interface', 'Capability', 'Neighbor Interface']
        header_format = workbook.add_format({'bold': True, 'bottom': 2})

        # Write column headers
        sheet.write_row(0, 0, fields, header_format)

        for j, val in enumerate(cat_details.lldp):
            sheet.write_row(j + 1, 0, list(val.values()))

    # Set workbook properties
    workbook.set_properties({'title': 'Catalyst Client Details'})

    # Close workbook and return Excel file as a response to the request
    workbook.close()

    console.print(f"[green]Download complete![/]")

    # Return Excel file as a response to the request
    response = Response(output.getvalue(), mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    response.headers.set('Content-Disposition', 'attachment', filename='catalyst_client_details.xlsx')
    return response


@app.route('/download/client/meraki')
def download_meraki_client():
    """
    Download Excel file containing meraki client information seen in WebGUI Tables (event triggered by download button)
    :return: Excel File
    """
    console.print(f"Downloading Excel File [green]meraki_client_details.xlsx[/]")

    # Create in-memory file for writing Excel data
    output = BytesIO()

    # Create Excel workbook and add sheets
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    sheets = []

    # Add Network Sheets
    for name in meraki_details.sorted_net_names:
        sheet = workbook.add_worksheet(name)
        sheets.append(sheet)

    # Define Column Headers (Meraki Client Details)
    fields = ['Field', 'Value']
    header_format = workbook.add_format({'bold': True, 'bottom': 2})

    for sheet in sheets:
        target_dict = {}

        # Get Network Sheet
        for net in meraki_details.clientDetails['networks']:
            if net['network_name'] == sheet.name:
                target_dict = net['client_details']

        # Write column headers
        sheet.write_row(0, 0, fields, header_format)

        # Write Column Rows
        if len(target_dict) > 0:
            sheet.write_row(1, 0, ['Status', target_dict['status'] if target_dict['status'] else 'None'])
            sheet.write_row(2, 0, ['Mac Address', target_dict['mac'] if target_dict['mac'] else 'None'])
            sheet.write_row(3, 0, ['IP Address', target_dict['ip'] if target_dict['ip'] else 'None'])
            sheet.write_row(4, 0, ['VLAN', target_dict['vlan'] if target_dict['vlan'] else 'None'])
            sheet.write_row(5, 0,
                            ['Device Manufacturer',
                             target_dict['manufacturer'] if target_dict['manufacturer'] else 'None'])
            sheet.write_row(6, 0,
                            ['Device OS', target_dict['os'] if target_dict['os'] else 'None'])
            sheet.write_row(7, 0,
                            ['Device User',
                             target_dict['user'] if target_dict['user'] else 'None'])
            sheet.write_row(8, 0,
                            ['Device Description',
                             target_dict['description'] if target_dict['description'] else 'None'])
            sheet.write_row(9, 0,
                            ['Recent Device (Serial)',
                             target_dict['recentDeviceSerial'] if target_dict['recentDeviceSerial'] else 'None'])
            sheet.write_row(10, 0,
                            ['Recent Device (Name)',
                             target_dict['recentDeviceName'] if target_dict['recentDeviceName'] else 'None'])
            sheet.write_row(11, 0,
                            ['Connection Type', target_dict['recentDeviceConnection'] if target_dict[
                                'recentDeviceConnection'] else 'None'])

            if target_dict['recentDeviceConnection'] == 'Wired':
                sheet.write_row(12, 0,
                                ['Switchport', target_dict['switchport'] if target_dict[
                                    'switchport'] else 'None'])
            else:
                sheet.write_row(12, 0,
                                ['SSID', target_dict['ssid'] if target_dict[
                                    'ssid'] else 'None'])

    # Set workbook properties
    workbook.set_properties({'title': 'Meraki Client Details'})

    # Close workbook and return Excel file as a response to the request
    workbook.close()

    console.print(f"[green]Download complete![/]")

    # Return Excel file as a response to the request
    response = Response(output.getvalue(), mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    response.headers.set('Content-Disposition', 'attachment', filename='meraki_client_details.xlsx')
    return response


@app.route('/download/usage')
def download_usage():
    """
    Download Excel file containing app summary information seen in WebGUI Tables (event triggered by download button)
    :return: Excel File
    """
    console.print(f"Downloading Excel File [green]meraki_client_app_usage.xlsx[/]")

    # Create in-memory file for writing Excel data
    output = BytesIO()

    # Create Excel workbook and add sheets
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    sheets = []

    # Add Summary Sheet
    sheet = workbook.add_worksheet('Summary')
    sheets.append(sheet)

    # Add Network Sheets
    for name in meraki_details.sorted_net_names:
        sheet = workbook.add_worksheet(name)
        sheets.append(sheet)

    # Define Column Headers (Usage)
    fields = ['Application', 'Received', 'Sent']
    header_format = workbook.add_format({'bold': True, 'bottom': 2})

    for sheet in sheets:
        target_dict = {}

        # Write column headers
        sheet.write_row(0, 0, fields, header_format)

        # Write Summary Sheet
        if sheet.name == 'Summary':
            target_dict = meraki_details.usage['summary']
        # Write network sheet
        else:
            for net in meraki_details.usage['networks']:
                if net['network_name'] == sheet.name:
                    target_dict = net['applications']

        for j, (k, v_list) in enumerate(target_dict.items()):
            sheet.write_row(j + 1, 0, [k] + v_list)

    # Set workbook properties
    workbook.set_properties({'title': 'Meraki App Usage'})

    # Close workbook and return Excel file as a response to the request
    workbook.close()

    console.print(f"[green]Download complete![/]")

    # Return Excel file as a response to the request
    response = Response(output.getvalue(), mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    response.headers.set('Content-Disposition', 'attachment', filename='meraki_client_app_usage.xlsx')
    return response


if __name__ == '__main__':
    app.run(port=5000)
