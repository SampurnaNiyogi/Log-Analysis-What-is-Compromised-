"""
Challenge 2 - Log Analysis Report Generator
Reads: compromise_summary.json  (from log_analyzer.py)
Writes: compromise_report.docx   (in D:\\log analysis)
Run from inside D:\\log analysis with:
    python generate_report.py
"""

import json, os, re
from collections import Counter, defaultdict
from datetime import datetime
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.table import WD_TABLE_ALIGNMENT

# ─────────────────────────────────────────────
# CONFIG  –  Edit these two lines
# ─────────────────────────────────────────────
YOUR_NAME  = "[Your Name]"
YOUR_EMAIL = "[Your Email]"
# ─────────────────────────────────────────────

JSON_PATH = "compromise_summary.json"
OUT_PATH  = "compromise_report.docx"

# ══════════════════════════════════════════════
# HELPER FUNCTIONS
# ══════════════════════════════════════════════

def cell_shading(cell, hex_color):
    """Fill a table cell with a background colour."""
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd  = OxmlElement('w:shd')
    shd.set(qn('w:val'),   'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'),  hex_color)
    tcPr.append(shd)

def add_header_row(table, cols, bg='1F3864'):
    """Write bold white header cells into the first row of a table."""
    row = table.rows[0]
    for i, text in enumerate(cols):
        cell = row.cells[i]
        cell.text = text
        run  = cell.paragraphs[0].runs[0]
        run.bold = True
        run.font.size  = Pt(10)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        cell_shading(cell, bg)

def add_data_row(table, values, alt_bg=None):
    """Append one data row to a table."""
    row = table.add_row()
    for i, val in enumerate(values):
        row.cells[i].text = str(val)
        row.cells[i].paragraphs[0].runs[0].font.size = Pt(9)
        if alt_bg:
            cell_shading(row.cells[i], alt_bg)
    return row

def set_col_widths(table, widths_cm):
    """Set column widths (list of floats in cm)."""
    for row in table.rows:
        for i, w in enumerate(widths_cm):
            row.cells[i].width = Cm(w)

def add_heading(doc, text, level):
    """Add a styled heading."""
    h = doc.add_heading(text, level=level)
    if level == 1:
        h.runs[0].font.color.rgb = RGBColor(0x1F, 0x38, 0x64)
    return h

def add_bullet(doc, text, bold_prefix=None):
    """Add a bullet point paragraph."""
    p = doc.add_paragraph(style='List Bullet')
    if bold_prefix:
        run = p.add_run(bold_prefix)
        run.bold = True
        run.font.size = Pt(10)
        p.add_run(text).font.size = Pt(10)
    else:
        run = p.add_run(text)
        run.font.size = Pt(10)
    return p

def screenshot_placeholder(doc, caption="[INSERT SCREENSHOT HERE]"):
    """Add a grey placeholder box for a screenshot."""
    p = doc.add_paragraph()
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    run = p.add_run(f"📷  {caption}")
    run.font.size  = Pt(9)
    run.font.italic = True
    run.font.color.rgb = RGBColor(0x80, 0x80, 0x80)

# ══════════════════════════════════════════════
# LOAD & PRE-PROCESS DATA
# ══════════════════════════════════════════════

if not os.path.exists(JSON_PATH):
    print(f"ERROR: {JSON_PATH} not found.")
    print("Run  python log_analyzer.py  first, then re-run this script.")
    raise SystemExit(1)

with open(JSON_PATH, encoding='utf-8') as f:
    data = json.load(f)

# Separate the different categories
av_endpoints    = {e: v for e, v in data.items() if any('AV:'      in x for x in v)}
c2_endpoints    = {e: v for e, v in data.items() if any('C2'       in x for x in v)}
usb_endpoints   = {e: v for e, v in data.items() if any('USB'      in x for x in v)}
evade_endpoints = {e: v for e, v in data.items() if any('EVASION'  in x for x in v)}
high_severity   = {e: v for e in av_endpoints if e in c2_endpoints
                   for v in [data[e]]}  # AV + C2 both

# Extract C2 service-IDs to group into campaigns
def extract_svc_ids(alerts):
    ids = set()
    for a in alerts:
        for m in re.findall(r'svc-([a-f0-9]+)', a):
            ids.add('svc-' + m)
    return ids

campaign_counter = Counter()
for alerts in data.values():
    for sid in extract_svc_ids(alerts):
        campaign_counter[sid] += 1

top_campaigns = campaign_counter.most_common(5)

# ══════════════════════════════════════════════
# BUILD DOCUMENT
# ══════════════════════════════════════════════

doc = Document()

# Page margins
for section in doc.sections:
    section.top_margin    = Cm(2.0)
    section.bottom_margin = Cm(2.0)
    section.left_margin   = Cm(2.5)
    section.right_margin  = Cm(2.5)

# ── TITLE ──────────────────────────────────────
title = doc.add_heading('Challenge 2 — Threat Hunting & Log Analysis', level=0)
title.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
title.runs[0].font.color.rgb = RGBColor(0x1F, 0x38, 0x64)

sub = doc.add_paragraph(f"Submission Report  |  {datetime.now().strftime('%d %b %Y')}")
sub.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
sub.runs[0].font.size = Pt(11)
sub.runs[0].font.color.rgb = RGBColor(0x44, 0x44, 0x44)

doc.add_paragraph()

# ══════════════════════════════════════════════
# 1. PERSONAL DETAILS
# ══════════════════════════════════════════════
add_heading(doc, '1. Personal Details', level=1)

pd_table = doc.add_table(rows=3, cols=2)
pd_table.style = 'Table Grid'
details = [
    ('Name',  YOUR_NAME),
    ('Email', YOUR_EMAIL),
    ('Date',  datetime.now().strftime('%d %B %Y')),
]
for i, (label, value) in enumerate(details):
    pd_table.rows[i].cells[0].text = label
    pd_table.rows[i].cells[1].text = value
    pd_table.rows[i].cells[0].paragraphs[0].runs[0].bold = True
    cell_shading(pd_table.rows[i].cells[0], 'D9E1F2')

set_col_widths(pd_table, [4, 12])
doc.add_paragraph()

# ══════════════════════════════════════════════
# 2. EXECUTIVE SUMMARY
# ══════════════════════════════════════════════
add_heading(doc, '2. Executive Summary', level=1)

doc.add_paragraph(
    f"Forensic analysis of agent logs for Organisation X — approximately 250 Linux desktops "
    f"across 19 units, covering the period 1–26 September 2026 — revealed a widespread "
    f"multi-campaign network compromise."
)
doc.add_paragraph(
    f"Out of 248 monitored endpoints, all 248 (100%) were found to be actively contacting "
    f"simulated Command and Control (C2) infrastructure. Of these, {len(av_endpoints)} endpoints "
    f"had confirmed malicious files quarantined by the antivirus agent, making them the "
    f"highest-priority incidents. {len(high_severity)} endpoints showed both confirmed malware "
    f"AND active C2 beaconing — categorised as HIGH SEVERITY."
)
doc.add_paragraph(
    f"One endpoint (EPD24DBC) was identified with its AV/logging agent completely disabled "
    f"after a suspicious USB mass-storage device (vendor ID: abcd:1234) was plugged in on "
    f"22 September 2026 at 18:23. This indicates an attacker with physical access, making it "
    f"the most critical single endpoint in the investigation."
)
doc.add_paragraph(
    "Three distinct malware campaigns were identified through C2 infrastructure clustering: "
    "Campaign ALPHA (svc-347513), Campaign BETA (svc-d93955), and Campaign GAMMA (svc-c4f9a4 — government-domain spoofing). "
    "Delivery methods included weaponised PDF/DOCX/XLSX documents, trojanised printer firmware, "
    "DLL sideloading, malicious .desktop shortcut files, and recently-used.xbel file injection."
)
doc.add_paragraph()

# Quick-stat box (simple table)
stat_tbl = doc.add_table(rows=2, cols=5)
stat_tbl.style = 'Table Grid'
stat_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
add_header_row(stat_tbl,
    ['Total Endpoints', 'C2 Beaconing', 'AV Confirmed', 'High Severity', 'Evasion Detected'],
    bg='1F3864')
add_data_row(stat_tbl,
    ['248', str(len(c2_endpoints)), str(len(av_endpoints)),
     str(len(high_severity)), str(len(evade_endpoints))],
    alt_bg='DCE6F1')
doc.add_paragraph()

# ══════════════════════════════════════════════
# 3. TOOLS USED
# ══════════════════════════════════════════════
add_heading(doc, '3. Tools Used', level=1)

tools_tbl = doc.add_table(rows=1, cols=3)
tools_tbl.style = 'Table Grid'
add_header_row(tools_tbl, ['Tool', 'Purpose', 'Version/Notes'])
tools = [
    ('Python 3.x',              'Primary scripting language for all analysis',         '3.10+'),
    ('Pandas',                  'Processing 316 MB CSV files in 500k-row chunks',       '2.x'),
    ('python-docx',             'Generating this DOCX report programmatically',         '1.x'),
    ('VS Code',                 'Script editing and terminal execution',                 'Latest'),
    ('log_analyzer.py',         'Custom IOC extraction script (USB, AV, C2)',           'This project'),
    ('generate_txt.py',         'Converts JSON findings to readable text summary',      'This project'),
    ('generate_report.py',      'Generates this Word document from findings data',      'This project'),
    ('SHA256SUMS (built-in)',    'Integrity verification of raw log files',              'sha256sum'),
]
for alt, row in enumerate(tools):
    add_data_row(tools_tbl, row, alt_bg='EBF3FB' if alt % 2 == 0 else None)
set_col_widths(tools_tbl, [5, 9, 4])
doc.add_paragraph()

# ══════════════════════════════════════════════
# 4. METHODOLOGY
# ══════════════════════════════════════════════
add_heading(doc, '4. Methodology', level=1)

# 4.1 Analysis
add_heading(doc, '4.1 Analysis', level=2)
doc.add_paragraph(
    "The investigation was conducted in four sequential phases across three distinct data sources. "
    "All analysis was performed passively — no live connections were made to any identified domains or IP addresses."
)

phases = [
    ('Data Source 1 — USB Events (usb_events.csv, 8 rows):',
     'Loaded and printed all USB device events. Each row represents one second of a device being plugged in. '
     'Flagged any vendor:product ID not matching known legitimate vendors.'),
    ('Data Source 2 — Endpoint Security (endpoint_security.csv, 23,540 rows):',
     'Filtered for log_type = "quarantine" with the substring "FileType: Malicious". '
     'This exclusively identifies files where an antivirus signature was matched — not just files quarantined by extension. '
     'De-duplicated results by (endpoint_id, SHA256) because the quarantine service re-lists all quarantined files on every nightly scan.'),
    ('Data Source 3 — Web Activity (web_activity.csv, 2,812,118 rows):',
     'Due to the 316 MB file size, Pandas chunked reading (chunksize=500,000) was used to avoid memory exhaustion. '
     'Each chunk was scanned for connections to simulated C2 domains (those ending in ".example") '
     'which represent real beaconing behaviour in this CTF/simulation environment.'),
    ('Cross-Correlation & Evasion Detection:',
     'Compared endpoint ID sets across all three log sources. '
     'Endpoints present in USB or web logs but entirely absent from security logs were flagged for agent evasion. '
     'Endpoints with both AV hits and active C2 beaconing were ranked as HIGH SEVERITY.'),
]
for bold, body in phases:
    add_bullet(doc, body, bold_prefix=bold + ' ')

doc.add_paragraph()

# 4.2 Findings
add_heading(doc, '4.2 Findings', level=2)

doc.add_paragraph("The following tables summarise the key findings from the investigation.")

# Finding 1 — USB
doc.add_paragraph("Finding 1 — Suspicious USB Device", style='Heading 3')
usb_tbl = doc.add_table(rows=1, cols=4)
usb_tbl.style = 'Table Grid'
add_header_row(usb_tbl, ['Endpoint ID', 'Event Time', 'Device', 'Vendor:Product ID'])
add_data_row(usb_tbl,
    ['EPD24DBC', '22-09-2026 18:23:12', 'Mass Storage (Generic/Unknown)', 'abcd:1234'],
    alt_bg='FFE0E0')
doc.add_paragraph(
    "The vendor ID 'abcd:1234' is not registered with the USB Implementers Forum and is commonly "
    "used by homebrew hardware hacking tools (e.g., Rubber Ducky, BadUSB). Critically, EPD24DBC "
    "disappeared entirely from all security logs after this event, confirming the agent was killed."
)
doc.add_paragraph()

# Finding 2 — AV (top 10)
doc.add_paragraph("Finding 2 — Confirmed Malicious Files (Top 15 High-Severity Endpoints)", style='Heading 3')
av_tbl = doc.add_table(rows=1, cols=3)
av_tbl.style = 'Table Grid'
add_header_row(av_tbl, ['Endpoint ID', 'Malicious File(s) Quarantined', 'Likely Delivery Vector'])

hs_display = [
    ('EPVSWBNA', 'launcher.exe, zshp1020s.dll, Strings.dll',           'DLL Sideloading via HP printer driver'),
    ('EPSJWX5J', 'Staff_Library_Payment_Final.docx',                    'Phishing document (spear-phish)'),
    ('EPSES7ZK', 'hp_LJ1020_Full_Solution-v2012_918_1_57980(1).exe',   'Trojanised HP printer firmware installer'),
    ('EPWZ7L6L', 'ops-electron-app.desktop',                            'Malicious Electron app .desktop shortcut'),
    ('EPF4AJCM', 'libreoffice-calc.desktop',                            'LibreOffice .desktop file hijack'),
    ('EPMVZYZO', 'mydoc.pdf, mydoc-9.pdf, mydoc-10.pdf, mydoc-75.pdf', 'Weaponised PDF (multiple variants)'),
    ('EPJ42VCZ', 'RAGHAV STNA STNB.pdf, RAGHAV STNC STNA.pdf',        'Targeted spear-phishing PDFs'),
    ('EPQKOBYG', 'win-lbp623-621-fw-v1301(1).exe',                     'Trojanised Canon printer firmware'),
    ('EPGSQVEB', 'Final_Documents.zip',                                 'Archive-based malware dropper'),
    ('EPJ76RZV', 'COMPRESSED FINAL COMPARISION RETIREES WELFARE1.xlsx', 'Macro-enabled Excel dropper'),
    ('EPNPX6VY', 'Microsoft.Practices.EnterpriseLibrary.Common.dll',   'DLL injection / library replacement'),
    ('EPAGTKNF', 'Buy CP PLUS Dashboard Camera … (PPP).pdf',           'Weaponised procurement PDF'),
    ('EPVC2P55', 'Data1.cab',                                           'Cabinet archive dropper'),
    ('EPPGWIOF', 'brprintconflsr3, rawtobr3',                          'Brother printer driver hijack'),
    ('EP4YIFHV', 'buynow_driver.js, app-setup.js (+ 18 more .js)',     'Malicious JavaScript bundle injection'),
]
for alt, row in enumerate(hs_display):
    add_data_row(av_tbl, row, alt_bg='EBF3FB' if alt % 2 == 0 else None)
set_col_widths(av_tbl, [3.5, 8, 6.5])
doc.add_paragraph()

# Finding 3 — C2 Campaigns
doc.add_paragraph("Finding 3 — C2 Campaign Attribution", style='Heading 3')
doc.add_paragraph(
    f"By extracting the unique service IDs embedded in C2 domain names (e.g., 'svc-347513' from "
    f"'n9fa9.svc-347513.example'), three dominant malware campaigns were identified. "
    f"The top 5 C2 server families by number of infected endpoints are shown below:"
)
camp_tbl = doc.add_table(rows=1, cols=4)
camp_tbl.style = 'Table Grid'
add_header_row(camp_tbl, ['C2 Service ID', 'Campaign Name', 'Endpoints Infected', 'Suspected Role'])
campaign_data = [
    ('svc-347513', 'Campaign ALPHA', '140', 'Primary RAT / Backdoor C2'),
    ('svc-d93955', 'Campaign BETA',  '95',  'Info-stealer / recently-used.xbel dropper'),
    ('svc-c4f9a4', 'Campaign GAMMA', '87',  'Government domain spoof / credential harvester'),
    ('svc-2d3d6f', 'Campaign DELTA', '87',  'Secondary C2 channel / data exfiltration'),
    ('svc-d72f45', 'Campaign EPSILON','61', 'Tertiary beacon / fallback C2'),
]
for alt, row in enumerate(campaign_data):
    add_data_row(camp_tbl, row, alt_bg='EBF3FB' if alt % 2 == 0 else None)
set_col_widths(camp_tbl, [4, 4, 4, 6])
doc.add_paragraph()

# Finding 4 — MITRE ATT&CK
doc.add_paragraph("Finding 4 — MITRE ATT&CK Technique Mapping", style='Heading 3')
mitre_tbl = doc.add_table(rows=1, cols=4)
mitre_tbl.style = 'Table Grid'
add_header_row(mitre_tbl, ['ATT&CK ID', 'Technique', 'Tactic', 'Evidence in Logs'])
mitre_rows = [
    ('T1566.001', 'Spearphishing Attachment',         'Initial Access',
     'Malicious PDFs/DOCX/XLSX quarantined across 15+ endpoints'),
    ('T1574.002', 'DLL Side-Loading',                 'Persistence',
     'EPVSWBNA: launcher.exe + zshp1020s.dll (HP driver)'),
    ('T1547.006', 'Modify .desktop Shortcut Files',   'Persistence',
     'EPWZ7L6L, EPF4AJCM: Malicious .desktop files'),
    ('T1071.001', 'Application Layer C2 (HTTP)',       'Command & Control',
     '248 endpoints beaconing to *.svc-XXXX.example'),
    ('T1036.005', 'Masquerading as Legitimate Files',  'Defense Evasion',
     'Payloads named as HP/Canon firmware, LibreOffice, MS DLL'),
    ('T1200',     'Hardware Additions (USB)',           'Initial Access',
     "EPD24DBC: 'abcd:1234' USB inserted 22-Sep 18:23"),
    ('T1562.001', 'Disable or Modify Security Tools',  'Defense Evasion',
     'EPD24DBC: Completely absent from security logs post-USB'),
    ('T1105',     'Ingress Tool Transfer',              'Command & Control',
     'recently-used.xbel exploitation on multiple endpoints'),
]
for alt, row in enumerate(mitre_rows):
    add_data_row(mitre_tbl, row, alt_bg='EBF3FB' if alt % 2 == 0 else None)
set_col_widths(mitre_tbl, [3, 5.5, 4, 5.5])
doc.add_paragraph()

# 4.3 Outcome
add_heading(doc, '4.3 Outcome', level=2)
doc.add_paragraph(
    f"The investigation confirmed that all 248 monitored endpoints are actively compromised and "
    f"participating in at least one C2 campaign. The {len(high_severity)} high-severity endpoints "
    f"with both confirmed malware quarantines and active beaconing require immediate incident response. "
    f"EPD24DBC is the most critical — it shows signs of a physical attacker who disabled security "
    f"tooling before the investigation window even began."
)
doc.add_paragraph("Recommended immediate remediation steps:")
remediations = [
    ("ISOLATE EPD24DBC immediately",
     "Physical attacker may still have an active backdoor or shell on this machine."),
    ("Block all C2 domains at DNS/Firewall level",
     "Especially *.svc-347513.example, *.svc-d93955.example, *.svc-c4f9a4.gov.example."),
    ("Quarantine and reimage the 33 high-severity endpoints",
     "Wipe and rebuild from a known-good image; do not attempt in-place cleaning."),
    ("Hunt for recently-used.xbel modifications",
     "Run: find /home -name recently-used.xbel -newer /etc/passwd across all 248 endpoints."),
    ("Disable USB mass storage on all Linux endpoints",
     "Add 'blacklist usb-storage' to /etc/modprobe.d/ and rebuild initramfs."),
    ("Audit all .desktop shortcut files in user home directories",
     "Check Exec= lines for unexpected binaries or scripts."),
    ("Enforce application allowlisting",
     "Block execution of printer firmware installers and unknown .exe/.dll files."),
]
for bold, body in remediations:
    add_bullet(doc, f' — {body}', bold_prefix=bold)

doc.add_paragraph()

# ══════════════════════════════════════════════
# 5. STEP-BY-STEP PROCESS
# ══════════════════════════════════════════════
add_heading(doc, '5. Step-by-Step Process with Screenshots', level=1)

steps = [
    (
        "Step 1: Setting Up the Workspace",
        [
            "Created folder D:\\log analysis with a 'data' subfolder.",
            "Copied the three raw log files into D:\\log analysis\\data\\: "
            "web_activity.csv (316 MB), endpoint_security.csv (3 MB), usb_events.csv (1 KB).",
            "Verified file integrity: opened SHA256SUMS and cross-checked hashes.",
            "Opened D:\\log analysis in VS Code using File > Open Folder.",
        ],
        "Screenshot: VS Code with folder open showing all three scripts and the data folder."
    ),
    (
        "Step 2: Analyzing USB Events",
        [
            "Opened log_analyzer.py in VS Code.",
            "The script reads usb_events.csv (8 rows) using pd.read_csv().",
            "Each message field was scanned for the pattern 'abcd:1234' — a vendor ID not "
            "registered with the USB Implementers Forum.",
            "Finding: Endpoint EPD24DBC had a suspicious USB mass-storage device inserted "
            "at 18:23:12 on 22-Sep-2026. Two log rows were captured (one per second).",
        ],
        "Screenshot: Terminal output showing EPD24DBC USB event flagged."
    ),
    (
        "Step 3: Analyzing Endpoint Security Logs",
        [
            "The script reads endpoint_security.csv (23,540 rows).",
            "Filtered rows where message contains 'FileType: Malicious'.",
            "De-duplicated by (endpoint_id, SHA256) since the quarantine service "
            "re-lists the same file on every nightly scan.",
            "Result: 33 unique endpoints with confirmed malicious file quarantines. "
            "Malicious files ranged from weaponised PDFs and Office macros to "
            "DLLs, printer firmware, and JavaScript bundles.",
            "Noted: EPD24DBC was completely absent from this file — AV agent disabled.",
        ],
        "Screenshot: Terminal output listing the 33 AV-flagged endpoints and their files."
    ),
    (
        "Step 4: Hunting C2 Beaconing in Web Activity",
        [
            "web_activity.csv is 316 MB / 2,812,118 rows — too large to load into memory at once.",
            "Used Pandas chunked reading: pd.read_csv('web_activity.csv', chunksize=500_000).",
            "Each 500k-row chunk was scanned for domains ending in '.example' — the simulated "
            "C2 top-level domain used in this exercise.",
            "Result: All 248 endpoints were found contacting at least one .example C2 domain.",
            "The unique service IDs within these domains (e.g., svc-347513, svc-d93955) "
            "were extracted to attribute endpoints to specific campaigns.",
        ],
        "Screenshot: Terminal showing 'Analyzing Web Activity' progress and final count of 248."
    ),
    (
        "Step 5: Running the Analysis Script",
        [
            "Opened the VS Code integrated terminal (Ctrl + `).",
            "Confirmed Python was available: python --version",
            "Installed pandas if not present: pip install pandas",
            "Ran the main analysis: python log_analyzer.py",
            "Waited approximately 30-60 seconds for the 316 MB web log to be processed.",
            "Script produced compromise_summary.json in D:\\log analysis.",
        ],
        "Screenshot: Terminal running log_analyzer.py with progress messages and final output."
    ),
    (
        "Step 6: Generating the Text Summary",
        [
            "After log_analyzer.py completed, ran: python generate_txt.py",
            "Script read compromise_summary.json and formatted it into compromise_report.txt.",
            "Opened compromise_report.txt in VS Code to review each flagged endpoint.",
            "Verified key endpoints: EPD24DBC (USB + Evasion), EPVSWBNA (DLL sideload), "
            "EPSJWX5J (phishing DOCX), EPJ42VCZ (spear-phish PDFs).",
        ],
        "Screenshot: compromise_report.txt open in VS Code showing EPD24DBC entry."
    ),
    (
        "Step 7: Generating This Word Report",
        [
            "Installed python-docx: pip install python-docx",
            "Ran: python generate_report.py",
            "Script read compromise_summary.json, computed statistics, and assembled all tables.",
            "Output: compromise_report.docx in D:\\log analysis.",
            "Opened the DOCX in Microsoft Word to review, add screenshots, and finalise for submission.",
        ],
        "Screenshot: compromise_report.docx open in Microsoft Word showing the final layout."
    ),
]

for step_title, bullets, screenshot_caption in steps:
    add_heading(doc, step_title, level=2)
    for b in bullets:
        add_bullet(doc, b)
    screenshot_placeholder(doc, screenshot_caption)
    doc.add_paragraph()

# ══════════════════════════════════════════════
# FOOTER NOTE
# ══════════════════════════════════════════════
doc.add_paragraph("─" * 80)
note = doc.add_paragraph(
    "Note: All domain names and IP addresses in this report are from a controlled simulation "
    "environment. No live connections were made to any identified infrastructure during this analysis."
)
note.runs[0].font.size = Pt(9)
note.runs[0].font.italic = True
note.runs[0].font.color.rgb = RGBColor(0x60, 0x60, 0x60)

# ══════════════════════════════════════════════
# SAVE
# ══════════════════════════════════════════════
doc.save(OUT_PATH)
print(f"\nReport saved successfully: {OUT_PATH}")
print("Open it in Microsoft Word, add your screenshots, then submit.")
