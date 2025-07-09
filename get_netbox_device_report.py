#!/usr/bin/env python3

import os
import sys
import requests
from collections import defaultdict
from datetime import datetime
import io

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

if not NETBOX_URL or not NETBOX_TOKEN:
    print("Missing NETBOX_URL (or NETBOX_API) or NETBOX_TOKEN environment variables.", file=sys.stderr)
    sys.exit(1)

headers = {
    'Authorization': f"Token {NETBOX_TOKEN}",
    'Accept': 'application/json',
}

# ---- Device Role Grouping Mapping ----
DEVICE_ROLE_GROUPS = {
    "Switches": {
        "Enterprise Core": [5],
        "Enterprise Edge": [1],
        "Operational": [28],
    },
    "Gateway/Router": {
        "Enterprise": [12],
        "Operational": [34],
    },
    "UPS": {
        "Enterprise": [14],
        "Operational": [44],
    },
    "Physical Hosts": {
        "Enterprise": [6],
        "Operational": [43],
    },
    "SAN": {
        "Enterprise": [16],
    },
    "NAS": {
        "Enterprise": [15],
        "Operational": [41],
    },
    "VM": {
        "Enterprise": [4, 19, 17, 20, 18, 33],
        "Operational": [35, 36, 37, 38, 39, 40],
    },
    "Operational Workstations": {
        "Operational": [30],
        "Quality Assurance": [32],
        "Optimisation": [31],
    },
}

# ---- Helper Functions ----
def get_heading_and_subheading(role_id):
    for heading, sub_map in DEVICE_ROLE_GROUPS.items():
        for subheading, role_ids in sub_map.items():
            if role_id in role_ids:
                return heading, subheading
    return "Other", "Other"

# ---- Main Data Gathering ----
site_device_counts = defaultdict(lambda: defaultdict(lambda: defaultdict(int)))

url = f"{NETBOX_URL.rstrip('/')}/api/dcim/devices/?limit=1000&expand=device_role,site"

while url:
    r = requests.get(url, headers=headers, timeout=30)
    r.raise_for_status()
    result = r.json()
    for device in result.get('results', []):
        site = device.get('site', {}).get('name', 'Unassigned Site')
        role_id = device.get('device_role', {}).get('id', None)
        if role_id is not None:
            heading, subheading = get_heading_and_subheading(role_id)
            site_device_counts[site][heading][subheading] += 1
        else:
            site_device_counts[site]['Other']['Other'] += 1
    url = result.get('next')

# ---- PDF Generation ----
pdf_buffer = io.BytesIO()
doc = SimpleDocTemplate(pdf_buffer, pagesize=letter)
styles = getSampleStyleSheet()
story = []

story.append(Paragraph("NetBox Device Count by Site, Heading, and Subheading", styles['Title']))
story.append(Spacer(1, 24))
date_str = datetime.now().strftime("%B %d, %Y")
story.append(Paragraph(f"Generated on: {date_str}", styles['Normal']))
story.append(Spacer(1, 24))

total_devices = 0
for site in sorted(site_device_counts):
    story.append(Paragraph(f"Site: {site}", styles['Heading2']))
    for heading in DEVICE_ROLE_GROUPS:
        if heading not in site_device_counts[site]:
            continue
        story.append(Paragraph(f"{heading}", styles['Heading3']))
        bullet_points = []
        for subheading in DEVICE_ROLE_GROUPS[heading]:
            count = site_device_counts[site][heading].get(subheading, 0)
            if count > 0:
                bullet_points.append(ListItem(Paragraph(f"{subheading}: {count}", styles['Normal'])))
                total_devices += count
        if bullet_points:
            story.append(ListFlowable(bullet_points, bulletType='bullet'))
        story.append(Spacer(1, 8))
    # Show "Other" category if present
    if "Other" in site_device_counts[site] and site_device_counts[site]["Other"]["Other"] > 0:
        story.append(Paragraph("Other: {}".format(site_device_counts[site]["Other"]["Other"]), styles['Normal']))
    story.append(Spacer(1, 12))

story.append(Spacer(1, 24))
story.append(Paragraph(f"<b>Total Devices in All Sites: {total_devices}</b>", styles['Normal']))

doc.build(story)
pdf_buffer.seek(0)
with open("/runner/netbox_device_report.pdf", "wb") as f:
    f.write(pdf_buffer.read())
print("Report generated: /runner/netbox_device_report.pdf")
