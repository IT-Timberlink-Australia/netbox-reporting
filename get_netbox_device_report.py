#!/usr/bin/env python3

import os
import sys
import requests
import json
from collections import defaultdict

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
except ImportError:
    print("openpyxl is not installed. Run 'pip install openpyxl' and retry.", file=sys.stderr)
    sys.exit(1)

# ---- Configuration ----
NETBOX_URL = os.environ.get('NETBOX_URL') or os.environ.get('NETBOX_API')
NETBOX_TOKEN = os.environ.get('NETBOX_TOKEN')
EXCLUDED_ROLE_IDS = {2, 11}

if not NETBOX_URL or not NETBOX_TOKEN:
    print("Missing NETBOX_URL (or NETBOX_API) or NETBOX_TOKEN environment variables.", file=sys.stderr)
    sys.exit(1)

headers = {
    'Authorization': f"Token {NETBOX_TOKEN}",
    'Accept': 'application/json',
}

HEADING_ORDER = [
    "Gateway/Router",
    "Switches",
    "Physical Hosts",
    "Virtual Machines",
    "Storage Area Network",
    "Network Attached Storage",
    "Production Workstations",
    "Uninterruptible Power Supply",
    "Printers",
    "Wireless Access Equipment",
]

DEVICE_ROLE_GROUPS = {
    "Gateway/Router": {
        "Enterprise": [12],
        "Operational": [34],
    },
    "Switches": {
        "Enterprise Core": [5],
        "Enterprise Edge": [1],
        "Operational": [28],
    },
    "Physical Hosts": {
        "Enterprise": [6],
        "Operational": [43],
    },
    "Virtual Machines": {
        "Enterprise": [4, 19, 17, 20, 18, 33],
        "Operational": [35, 36, 37, 38, 39, 40],
    },
    "Storage Area Network": {
        "Enterprise": [16],
    },
    "Network Attached Storage": {
        "Enterprise": [15],
        "Operational": [41],
    },
    "Production Workstations": {
        "Operational": [30],
        "Quality Assurance": [32],
        "Optimisation": [31],
    },
    "Uninterruptible Power Supply": {
        "Enterprise": [14],
        "Operational": [44],
    },
    "Printers": {
        "Enterprise": [24],
        "Operational": [26],
    },
    "Wireless Access Equipment": {
        "Access Points": [10],
        "Point to Point": [45],
        "Access Equipment": [46],
    },
}

def get_heading_and_subheading(role_id):
    for heading, sub_map in DEVICE_ROLE_GROUPS.items():
        for subheading, role_ids in sub_map.items():
            if role_id in role_ids:
                return heading, subheading
    return "Other", "Other"

def tick(val):
    return "✓" if val else "✗"

def short_desc(desc, length=100):
    if desc and len(desc) > length:
        return desc[:length-3] + "..."
    return desc or ""

# ---- Data Gathering ----
site_device_counts = defaultdict(lambda: defaultdict(lambda: defaultdict(lambda: {'count': 0, 'devices': []})))
device_debug_file = "/runner/device_debug.json"
with open(device_debug_file, "w") as dbg:
    dbg.write("")

def fetch_netbox_items(url, item_type):
    all_items = []
    while url:
        r = requests.get(url, headers=headers, timeout=30)
        r.raise_for_status()
        result = r.json()
        for item in result.get('results', []):
            item['_item_type'] = item_type
            all_items.append(item)
        url = result.get('next')
    return all_items

devices_url = f"{NETBOX_URL.rstrip('/')}/api/dcim/devices/?limit=1000&expand=role,site"
devices = fetch_netbox_items(devices_url, "device")

vms_url = f"{NETBOX_URL.rstrip('/')}/api/virtualization/virtual-machines/?limit=1000&expand=role,site"
vms = fetch_netbox_items(vms_url, "vm")

all_items = devices + vms

for item in all_items:
    with open(device_debug_file, "a") as dbg:
        dbg.write(json.dumps(item, indent=2) + "\n\n")
    site = item.get('site', {}).get('name', 'Unassigned Site')
    status = item.get('status', {}).get('value') if isinstance(item.get('status'), dict) else item.get('status')
    if status not in ('active', 1):
        continue
    role_id = None
    item_role = item.get('role', None)
    if isinstance(item_role, dict):
        role_id = item_role.get('id', None)
    elif isinstance(item_role, int):
        role_id = item_role
    else:
        role_id = None
    if role_id in EXCLUDED_ROLE_IDS:
        continue
    cf = item.get('custom_fields', {}) or {}
    device_info = {
        'site': site,
        'heading': None,
        'subheading': None,
        'name': item.get('name', 'Unknown Device'),
        'description': item.get('description', ''),
        'primary_ip': item.get('primary_ip'),
        'serial': item.get('serial'),
        'backup_primary': cf.get('last_backup_data_prim'),
        'monitoring_required': cf.get('mon_required'),
    }
    if role_id is not None:
        heading, subheading = get_heading_and_subheading(role_id)
        device_info['heading'] = heading
        device_info['subheading'] = subheading
        site_device_counts[site][heading][subheading]['count'] += 1
        site_device_counts[site][heading][subheading]['devices'].append(device_info)
    else:
        device_info['heading'] = "Other"
        device_info['subheading'] = "Other"
        site_device_counts[site]['Other']['Other']['count'] += 1
        site_device_counts[site]['Other']['Other']['devices'].append(device_info)

# ---- Excel Generation ----
wb = openpyxl.Workbook()
default_ws = wb.active
default_ws.title = "All Sites"

headers = [
    "Site", "Heading", "Subheading", "Device Name", "Description", "Primary IP",
    "Serial", "Backup Data - Primay", "Monitoring Required"
]
default_ws.append(headers)

header_fill = PatternFill("solid", fgColor="00336699")
header_font = Font(bold=True, color="FFFFFFFF")
for cell in default_ws[1]:
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = Alignment(horizontal="center", vertical="center")

# Write all devices as flat rows to 'All Sites' worksheet
for site in sorted(site_device_counts):
    for heading in HEADING_ORDER:
        if heading not in site_device_counts[site]:
            continue
        for subheading in DEVICE_ROLE_GROUPS[heading]:
            devices = site_device_counts[site][heading][subheading]['devices']
            for device in sorted(devices, key=lambda d: d['name']):
                default_ws.append([
                    site,
                    heading,
                    subheading,
                    device['name'],
                    short_desc(device.get('description', ''), 100),
                    tick(device.get('primary_ip')),
                    tick(device.get('serial')),
                    tick(device.get('backup_primary')),
                    tick(device.get('monitoring_required') is not False),
                ])

# Optional: auto-size columns in the default worksheet
for col in default_ws.columns:
    max_length = max(len(str(cell.value) or "") for cell in col)
    default_ws.column_dimensions[col[0].column_letter].width = min(max_length + 4, 50)

# ---- Create a worksheet for each site ----
for site in sorted(site_device_counts):
    ws = wb.create_sheet(title=site[:31])  # Excel worksheet name max length is 31
    ws.append(headers)
    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for heading in HEADING_ORDER:
        if heading not in site_device_counts[site]:
            continue
        for subheading in DEVICE_ROLE_GROUPS[heading]:
            devices = site_device_counts[site][heading][subheading]['devices']
            for device in sorted(devices, key=lambda d: d['name']):
                ws.append([
                    site,
                    heading,
                    subheading,
                    device['name'],
                    short_desc(device.get('description', ''), 100),
                    tick(device.get('primary_ip')),
                    tick(device.get('serial')),
                    tick(device.get('backup_primary')),
                    tick(device.get('monitoring_required') is not False),
                ])
    # Auto-size columns in each site worksheet
    for col in ws.columns:
        max_length = max(len(str(cell.value) or "") for cell in col)
        ws.column_dimensions[col[0].column_letter].width = min(max_length + 4, 50)

# Save file
excel_file = "/runner/netbox_device_report.xlsx"
wb.save(excel_file)
print("Excel report generated:", excel_file)
