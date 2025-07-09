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

# --- Config ---
NETBOX_URL = os.environ.get('NETBOX_API')
NETBOX_TOKEN = os.environ.get('NETBOX_TOKEN')

if not NETBOX_URL or not NETBOX_TOKEN:
    print("Missing NETBOX_URL or NETBOX_TOKEN environment variables.", file=sys.stderr)
    sys.exit(1)

headers = {
    'Authorization': f"Token {NETBOX_TOKEN}",
    'Accept': 'application/json',
}

# --- Gather data ---
site_device_counts = defaultdict(lambda: defaultdict(int))
url = f"{NETBOX_URL.rstrip('/')}/api/dcim/devices/?limit=1000&expand=device_role"
while url:
    r = requests.get(url, headers=headers, timeout=30)
    r.raise_for_status()
    result = r.json()
    for device in result.get('results', []):
        site = device.get('site', {}).get('name', 'Unassigned Site')
        # Debug: Dump a sample device to /tmp/device_debug.json
        with open("/tmp/device_debug.json", "w") as f:
            import json
            json.dump(device, f, indent=2)
        role = device.get('device_role', {}).get('name', 'Unknown Role')
        site_device_counts[site][role] += 1
    url = result.get('next')

# --- Generate PDF ---
pdf_buffer = io.BytesIO()
doc = SimpleDocTemplate(pdf_buffer, pagesize=letter)
styles = getSampleStyleSheet()
story = []

story.append(Paragraph("NetBox Device Count by Site and Role", styles['Title']))
story.append(Spacer(1, 24))
date_str = datetime.now().strftime("%B %d, %Y")
story.append(Paragraph(f"Generated on: {date_str}", styles['Normal']))
story.append(Spacer(1, 24))

total_devices = 0
for site in sorted(site_device_counts):
    story.append(Paragraph(f"Site: {site}", styles['Heading2']))
    bullet_points = []
    for dtype, count in sorted(site_device_counts[site].items()):
        bullet_points.append(ListItem(Paragraph(f"{dtype}: {count}", styles['Normal'])))
        total_devices += count
    story.append(ListFlowable(bullet_points, bulletType='bullet'))
    story.append(Spacer(1, 12))

story.append(Spacer(1, 24))
story.append(Paragraph(f"<b>Total Devices in All Sites: {total_devices}</b>", styles['Normal']))

doc.build(story)
pdf_buffer.seek(0)
with open("netbox_device_report.pdf", "wb") as f:
    f.write(pdf_buffer.read())
print("Report generated: netbox_device_report.pdf")
