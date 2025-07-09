#!/usr/bin/env python3

import os
import sys
import requests
from collections import defaultdict
from datetime import datetime
import io
import json

try:
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, ListFlowable, ListItem, PageBreak
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib.pagesizes import letter
except ImportError:
    print("reportlab is not installed. Run 'pip install reportlab' and retry.", file=sys.stderr)
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

def tick(value):
    return "✓" if value else "✗"

# ---- Main Data Gathering ----
site_device_counts = defaultdict(lambda: defaultdict(lambda: defaultdict(lambda: {'count': 0, 'names': []})))
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
            item['_item_type'] = item_type  # device or vm
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

    # Only "active" status
    status = item.get('status', {}).get('value') if isinstance(item.get('status'), dict) else item.get('status')
    if status not in ('active', 1):
        continue

    # Role logic
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

    # --- Gather device info for the list ---
    cf = item.get('custom_fields', {}) or {}
    device_info = {
        'name': item.get('name', 'Unknown Device'),
        'description': item.get('description', ''),
        'primary_ip': item.get('primary_ip'),
        'serial': item.get('serial'),
        # Adjust these custom field keys as needed!
        'backup_primary': cf.get('backup_data_primay'),
        'monitoring_required': cf.get('monitoring_required'),
    }

    if role_id is not None:
        heading, subheading = get_heading_and_subheading(role_id)
        site_device_counts[site][heading][subheading]['count'] += 1
        site_device_counts[site][heading][subheading]['names'].append(device_info)
    else:
        site_device_counts[site]['Other']['Other']['count'] += 1
        site_device_counts[site]['Other']['Other']['names'].append(device_info)

# ---- PDF Generation ----
pdf_buffer = io.BytesIO()
doc = SimpleDocTemplate(pdf_buffer, pagesize=letter)
styles = getSampleStyleSheet()
story = []

# Title
story.append(Paragraph("Timberlink CMDB Active Production Device Report", styles['Title']))
story.append(Spacer(1, 24))

# Table column titles
col_titles = (
    "<b>Name</b> | <b>Description</b> | <b>IP</b> | <b>Serial</b> | "
    "<b>Backup</b> | <b>Monitoring</b>"
)

total_devices = 0
site_list = sorted(site_device_counts)
for idx, site in enumerate(site_list):
    story.append(Paragraph(f"Site: {site}", styles['Heading2']))
    for heading in HEADING_ORDER:
        if heading not in site_device_counts[site]:
            continue
        story.append(Paragraph(f"{heading}", styles['Heading3']))
        bullet_points = []
        for subheading in DEVICE_ROLE_GROUPS[heading]:
            data = site_device_counts[site][heading].get(subheading, {'count': 0, 'names': []})
            count = data['count']
            if count > 0:
                bullet_points.append(ListItem(Paragraph(f"{subheading}: {count}", styles['Normal'])))
                # Add header for columns
                bullet_points.append(ListItem(Paragraph(col_titles, styles['Normal'])))
                # List device info, formatted as columns
                for device in sorted(data['names'], key=lambda d: d['name']):
                    description = device.get('description', '')
                    has_primary_ip = tick(device.get('primary_ip'))
                    has_serial = tick(device.get('serial'))
                    has_backup = tick(device.get('backup_primary'))
                    mon_req = device.get('monitoring_required')
                    monitoring_required = "✗" if mon_req is False else "✓"
                    device_line = (
                        f"<b>{device['name']}</b> | {description} | {has_primary_ip} | "
                        f"{has_serial} | {has_backup} | {monitoring_required}"
                    )
                    bullet_points.append(ListItem(Paragraph(device_line, styles['Normal'])))
                total_devices += count
        if bullet_points:
            story.append(ListFlowable(bullet_points, bulletType='bullet'))
        story.append(Spacer(1, 8))
    # Show "Other" category if present
    if "Other" in site_device_counts[site] and site_device_counts[site]["Other"]["Other"]['count'] > 0:
        count = site_device_counts[site]["Other"]["Other"]['count']
        names = site_device_counts[site]["Other"]["Other"]['names']
        story.append(Paragraph("Other: {}".format(count), styles['Normal']))
        # Add header for columns
        story.append(Paragraph(col_titles, styles['Normal']))
        for device in sorted(names, key=lambda d: d['name']):
            description = device.get('description', '')
            has_primary_ip = tick(device.get('primary_ip'))
            has_serial = tick(device.get('serial'))
            has_backup = tick(device.get('backup_primary'))
            mon_req = device.get('monitoring_required')
            monitoring_required = "✗" if mon_req is False else "✓"
            device_line = (
                f"<b>{device['name']}</b> | {description} | {has_primary_ip} | "
                f"{has_serial} | {has_backup} | {monitoring_required}"
            )
            story.append(Paragraph(device_line, styles['Normal']))
        story.append(Spacer(1, 8))
    story.append(Spacer(1, 12))
    if idx != len(site_list) - 1:
        story.append(PageBreak())

# Footer with date
story.append(Spacer(1, 12))
date_str = datetime.now().strftime("%B %d, %Y")
story.append(Paragraph(f"Generated on: {date_str}", styles['Normal']))

doc.build(story)
pdf_buffer.seek(0)
with open("/runner/netbox_device_report.pdf", "wb") as f:
    f.write(pdf_buffer.read())
print("Report generated: /runner/netbox_device_report.pdf")
