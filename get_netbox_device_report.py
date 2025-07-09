#!/usr/bin/env python3

import os
import sys
import requests
from collections import defaultdict
from datetime import datetime
import io
import json

try:
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, ListFlowable, ListItem
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib.pagesizes import letter
except ImportError:
    print("reportlab is not installed. Run 'pip install reportlab' and retry.", file=sys.stderr)
    sys.exit(1)

# ---- Configuration ----
NETBOX_URL = os.environ.get('NETBOX_URL') or os.environ.get('NETBOX_API')
NETBOX_TOKEN = os.environ.get('NETBOX_TOKEN')
EXCLUDED_ROLE_IDS = {2, 11}  # 24,26 are now INCLUDED as Printer!

if not NETBOX_URL or not NETBOX_TOKEN:
    print("Missing NETBOX_URL (or NETBOX_API) or NETBOX_TOKEN environment variables.", file=sys.stderr)
    sys.exit(1)

headers = {
    'Authorization': f"Token {NETBOX_TOKEN}",
    'Accept': 'application/json',
}

# ---- Heading Order ----
HEADING_ORDER = [
    "Gateway/Router",
    "Switches",
    "Physical Hosts",
    "Virtual Machines",
    "Storage Area Network",
    "Network Attached Storage",
    "Production Workstations",
    "Uninterruptible Power Supply",
    "Printer",
    "Wireless Access Equipment",
]

# ---- Device Role Grouping Mapping ----
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
    "Printer": {
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

# ---- Main Data Gathering ----
site_device_counts = defaultdict(lambda: defaultdict(lambda: defaultdict(lambda: {'count': 0, 'names': []})))
device_debug_file = "/runner/device_debug.json"
with open(device_debug_file, "w") as dbg:
    dbg.write("")

# Helper function to fetch paginated results from NetBox API
def fetch_netbox_items(url, item_type):
    all_items = []
    while url:
        r = requests.get(url, headers=headers, timeout=30)
        r.raise_for_status()
        result = r.json()
        for item in result.get('results', []):
            item['_item_type'] = item_type  # Track whether it's a device or VM
            all_items.append(item)
        url = result.get('next')
    return all_items

# Fetch Devices
devices_url = f"{NETBOX_URL.rstrip('/')}/api/dcim/devices/?limit=1000&expand=role,site"
devices = fetch_netbox_items(devices_url, "device")

# Fetch Virtual Machines
vms_url = f"{NETBOX_URL.rstrip('/')}/api/virtualization/virtual-machines/?limit=1000&expand=role,site"
vms = fetch_netbox_items(vms_url, "vm")

# Merge
all_items = devices + vms

for item in all_items:
    with open(device_debug_file, "a") as dbg:
        dbg.write(json.dumps(item, indent=2) + "\n\n")

    site = item.get('site', {}).get('name', 'Unassigned Site')

    # Filter: only process "active" devices/VMs
    status = item.get('status', {}).get('value') if isinstance(item.get('status'), dict) else item.get('status')
    if status not in ('active', 1):
        continue

    # Try to get role_id from dict or int
    role_id = None
    item_role = item.get('role', None)
    if isinstance(item_role, dict):
        role_id = item_role.get('id', None)
    elif isinstance(item_role, int):
        role_id = item_role
    else:
        role_id = None

    # Filter: skip excluded roles
    if role_id in EXCLUDED_ROLE_IDS:
        continue

    if role_id is not None:
        heading, subheading = get_heading_and_subheading(role_id)
        site_device_counts[site][heading][subheading]['count'] += 1
        site_device_counts[site][heading][subheading]['names'].append(item.get('name', 'Unknown Device'))
    else:
        site_device_counts[site]['Other']['Other']['count'] += 1
        site_device_counts[site]['Other']['Other']['names'].append(item.get('name', 'Unknown Device'))

# ---- PDF Generation ----
pdf_buffer = io.BytesIO()
doc = SimpleDocTemplate(pdf_buffer, pagesize=letter)
styles = getSampleStyleSheet()
story = []

# Title
story.append(Paragraph("Timberlink CMDB Active Production Device Report", styles['Title']))
story.append(Spacer(1, 24))

total_devices = 0
for site in sorted(site_device_counts):
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
                # List device names indented under the count
                name_bullets = [ListItem(Paragraph(name, styles['Normal'])) for name in sorted(data['names'])]
                if name_bullets:
                    bullet_points.append(ListFlowable(name_bullets, bulletType='bullet', leftIndent=18))
                total_devices += count
        if bullet_points:
            story.append(ListFlowable(bullet_points, bulletType='bullet'))
        story.append(Spacer(1, 8))
    # Show "Other" category if present
    if "Other" in site_device_counts[site] and site_device_counts[site]["Other"]["Other"]['count'] > 0:
        count = site_device_counts[site]["Other"]["Other"]['count']
        names = site_device_counts[site]["Other"]["Other"]['names']
        story.append(Paragraph("Other: {}".format(count), styles['Normal']))
        name_bullets = [ListItem(Paragraph(name, styles['Normal'])) for name in sorted(names)]
        if name_bullets:
            story.append(ListFlowable(name_bullets, bulletType='bullet', leftIndent=18))
    story.append(Spacer(1, 12))

story.append(Spacer(1, 24))
story.append(Paragraph(f"<b>Total Devices in All Sites: {total_devices}</b>", styles['Normal']))

# Footer with date
story.append(Spacer(1, 12))
date_str = datetime.now().strftime("%B %d, %Y")
story.append(Paragraph(f"Generated on: {date_str}", styles['Normal']))

doc.build(story)
pdf_buffer.seek(0)
with open("/runner/netbox_device_report.pdf", "wb") as f:
    f.write(pdf_buffer.read())
print("Report generated: /runner/netbox_device_report.pdf")
