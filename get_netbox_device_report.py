#!/usr/bin/env python3

import os
import sys
import requests
from collections import defaultdict
from datetime import datetime
import io
import json

try:
    from reportlab.platypus import (
        SimpleDocTemplate, Paragraph, Spacer, PageBreak,
        Table, TableStyle
    )
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.pagesizes import letter
    from reportlab.lib import colors
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

def color_tick(val):
    return '<font color="green">✓</font>' if val else '<font color="red">✗</font>'

def short_desc(desc, length=30):
    if desc and len(desc) > length:
        return desc[:length-3] + "..."
    return desc or ""

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
        'backup_primary': cf.get('last_backup_data_prim'),
        'monitoring_required': cf.get('mon_required'),
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
header_style = ParagraphStyle('Header', parent=styles['Normal'], alignment=1, fontSize=12, textColor=colors.darkblue, fontName="Helvetica-Bold")
subheader_style = ParagraphStyle('SubHeader', parent=styles['Normal'], alignment=0, fontSize=10, fontName="Helvetica-Bold")

story = []

# Title
story.append(Paragraph("Timberlink CMDB Active Production Device Report", header_style))
story.append(Spacer(1, 24))

total_devices = 0
site_list = sorted(site_device_counts)
for idx, site in enumerate(site_list):
    story.append(Paragraph(f"<b>Site: {site}</b>", header_style))
    for heading in HEADING_ORDER:
        if heading not in site_device_counts[site]:
            continue
        for subheading in DEVICE_ROLE_GROUPS[heading]:
            data = site_device_counts[site][heading].get(subheading, {'count': 0, 'names': []})
            if data['count'] > 0:
                story.append(Paragraph(f"{heading} - {subheading}: <b>{data['count']}</b>", subheader_style))
                # Build table data: headers + rows
                table_data = [[
                    "Device Name", "Description", "Primary IP", "Serial",
                    "Backup Data", "Monitoring"
                ]]
                for device in sorted(data['names'], key=lambda d: d['name']):
                    desc = short_desc(device.get('description', ''), 30)
                    table_data.append([
                        device.get('name', ''),
                        desc,
                        color_tick(device.get('primary_ip')),
                        color_tick(device.get('serial')),
                        color_tick(device.get('backup_primary')),
                        color_tick(device.get('monitoring_required') is not False),
                    ])
                # Create the table
                t = Table(table_data, repeatRows=1, hAlign='LEFT', colWidths=[90, 120, 55, 55, 70, 70])
                t.setStyle(TableStyle([
                    ('BACKGROUND', (0,0), (-1,0), colors.darkblue),
                    ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
                    ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                    ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0,0), (-1,-1), 8),
                    ('BOTTOMPADDING', (0,0), (-1,0), 6),
                    ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
                    ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.whitesmoke, colors.lightgrey]),
                ]))
                story.append(t)
                story.append(Spacer(1, 12))
                total_devices += data['count']
    if idx != len(site_list) - 1:
        story.append(PageBreak())

# Footer with date and total
story.append(Spacer(1, 24))
story.append(Paragraph(f"<b>Total Devices in All Sites: {total_devices}</b>", styles['Normal']))
story.append(Spacer(1, 12))
date_str = datetime.now().strftime("%B %d, %Y")
story.append(Paragraph(f"Generated on: {date_str}", styles['Normal']))

doc.build(story)
pdf_buffer.seek(0)
with open("/runner/netbox_device_report.pdf", "wb") as f:
    f.write(pdf_buffer.read())
print("Report generated: /runner/netbox_device_report.pdf")
