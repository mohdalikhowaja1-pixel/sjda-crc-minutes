import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import io

st.set_page_config(page_title="SJDA CRC Meeting Minutes Generator", layout="wide")

# Pre-defined Area Profiles with BLANK signatures
AREA_PROFILES = {
    "Tando Turail Council": {
        "region": "Southern",
        "venue": "Zoom Link / online",
        "staff": [
            {"name": "Amyn Hyder", "designation": "Program Manager", "sig": ""},
            {"name": "Salima", "designation": "Assistant Program Manager", "sig": ""},
            {"name": "Muhammad Ali Khowaja", "designation": "Family Mentor", "sig": ""}
        ],
        "crc": [
            {"name": "Dr. Sindhu Hajyani", "designation": "Chairperson", "sig": ""},
            {"name": "Aijaz Alwani", "designation": "Member", "sig": ""},
            {"name": "Razia Arshad", "designation": "Member", "sig": ""},
            {"name": "Nafisa Rahim", "designation": "Member", "sig": ""},
            {"name": "Salima Lakho", "designation": "Member", "sig": ""},
            {"name": "Seema Asif Molwani", "designation": "Member", "sig": ""},
            {"name": "Babar Ashraf", "designation": "Member", "sig": ""},
            {"name": "Shamshad Ali", "designation": "Member", "sig": ""}
        ]
    },
    "Hyderabad": {
        "region": "Southern",
        "venue": "Council Office - Hyderabad",
        "staff": [
            {"name": "Amyn Hyder", "designation": "Program Manager", "sig": ""},
            {"name": "Salima Khiyani", "designation": "Assistant Program Manager", "sig": ""},
            {"name": "Muhammad Ali Khowaja", "designation": "Family Mentor", "sig": ""}
        ],
        "crc": [
            {"name": "Kamran Karim Bux", "designation": "Chairperson", "sig": ""},
            {"name": "Hussain Khalil", "designation": "Member", "sig": ""},
            {"name": "Faheem Hussain", "designation": "Member", "sig": ""},
            {"name": "Noman Sultan", "designation": "Member", "sig": ""},
            {"name": "Farhana Kashif", "designation": "Member", "sig": ""},
            {"name": "Sanam Sikander", "designation": "Member", "sig": ""},
            {"name": "Arslan Khowaja", "designation": "Member", "sig": ""},
            {"name": "Muhammad Ali Feroz", "designation": "Member", "sig": ""}
        ]
    }
}

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    """Sets cell padding inside Word tables to match standard document styles."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for margin, value in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{margin}')
        node.set(qn('w:w'), str(value))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def format_cell(cell, text, bold=False, italic=False, size=9.5, align=WD_ALIGN_PARAGRAPH.LEFT):
    cell.text = text
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.space_before = Pt(2)
    for run in p.runs:
        run.font.name = 'Calibri'
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.italic = italic
    set_cell_margins(cell)

def generate_docx(location, meeting_date, meeting_time, venue, fdp_data, profile):
    doc = Document()
    
    # Configure document margins
    for section in doc.sections:
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.6)
        section.right_margin = Inches(0.6)

    # PAGE 1: MEETING INFO & FDP DECISIONS
    
    # Document Header
    p_hdr = doc.add_paragraph()
    p_hdr.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_hdr.paragraph_format.space_after = Pt(2)
    
    r1 = p_hdr.add_run("SILVER JUBILEE DEVELOPMENT AGENCY (SJDA)\n")
    r1.font.name = 'Calibri'
    r1.font.size = Pt(12)
    r1.font.bold = True
    
    r2 = p_hdr.add_run("Poverty Elimination Program (PE)\nCASE REVIEW COMMITTEE (CRC) MEETING MINUTES")
    r2.font.name = 'Calibri'
    r2.font.size = Pt(11)
    r2.font.bold = True

    # Section Header: Meeting Info
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run("MEETING INFORMATION")
    r.font.name = 'Calibri'
    r.font.size = Pt(10)
    r.font.bold = True

    info_table = doc.add_table(rows=6, cols=2)
    info_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    info_table.style = 'Table Grid'
    
    info_data = [
        ("Particular", "Details"),
        ("Region", profile["region"]),
        ("Location", location),
        ("Meeting Date", meeting_date.strftime("%d/%m/%Y")),
        ("Time", meeting_time),
        ("Venue / Zoom Link", venue)
    ]
    
    for i, (k, v) in enumerate(info_data):
        row = info_table.rows[i]
        format_cell(row.cells[0], k, bold=(i==0), size=9.5)
        format_cell(row.cells[1], v, bold=(i==0), size=9.5)
        row.cells[0].width = Inches(2.2)
        row.cells[1].width = Inches(4.8)

    # Section Header: Meeting Agenda
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run("MEETING AGENDA")
    r.font.name = 'Calibri'
    r.font.size = Pt(10)
    r.font.bold = True

    agenda_table = doc.add_table(rows=2, cols=1)
    agenda_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    agenda_table.style = 'Table Grid'
    format_cell(agenda_table.rows[0].cells[0], "Agenda Item", bold=True, size=9.5)
    format_cell(agenda_table.rows[1].cells[0], "CRC Review and Approval of Family Development Plans (FDPs)", size=9.5)
    agenda_table.rows[0].cells[0].width = Inches(7.0)
    agenda_table.rows[1].cells[0].width = Inches(7.0)

    # Section Header: Decisions on FDPs
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run("DECISIONS ON FDPS")
    r.font.name = 'Calibri'
    r.font.size = Pt(10)
    r.font.bold = True

    fdp_table = doc.add_table(rows=1, cols=8)
    fdp_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    fdp_table.style = 'Table Grid'
    
    headers = ["S. No.", "Beneficiary Name", "Family ID", "Poverty Level", "Total Economic Support (PKR)", "Loan (PKR)", "Grant (PKR)", "Status (Approved / Rejected / Deferred)"]
    col_widths = [Inches(0.5), Inches(1.8), Inches(0.8), Inches(0.8), Inches(0.9), Inches(0.7), Inches(0.7), Inches(1.1)]

    for i, h in enumerate(headers):
        format_cell(fdp_table.rows[0].cells[i], h, bold=True, size=8.5, align=WD_ALIGN_PARAGRAPH.CENTER)
        fdp_table.rows[0].cells[i].width = col_widths[i]

    for idx, item in enumerate(fdp_data, start=1):
        row_cells = fdp_table.add_row().cells
        vals = [
            str(idx),
            item["name"],
            item["family_id"],
            item["poverty_level"],
            f"{item['support']:,}" if isinstance(item['support'], (int, float)) and item['support'] > 0 else str(item['support']),
            f"{item['loan']:,}" if isinstance(item['loan'], (int, float)) and item['loan'] > 0 else str(item['loan']),
            f"{item['grant']:,}" if isinstance(item['grant'], (int, float)) and item['grant'] > 0 else str(item['grant']),
            item["status"]
        ]
        for i, val in enumerate(vals):
            align_style = WD_ALIGN_PARAGRAPH.CENTER if i in [0, 2, 3, 7] else WD_ALIGN_PARAGRAPH.LEFT
            format_cell(row_cells[i], val, size=8.5, align=align_style)
            row_cells[i].width = col_widths[i]

    # FORCE PAGE BREAK (Guarantees Page 2 is dedicated solely to Attendance & Signatures)
    doc.add_page_break()

    # PAGE 2: ATTENDANCE & SIGNATURES
    
    # Staff Attendance Section
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run("PE-SJDA STAFF ATTENDANCE")
    r.font.name = 'Calibri'
    r.font.size = Pt(10)
    r.font.bold = True

    staff_table = doc.add_table(rows=1, cols=4)
    staff_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    staff_table.style = 'Table Grid'
    staff_widths = [Inches(0.6), Inches(2.5), Inches(2.5), Inches(1.4)]
    
    for i, h in enumerate(["S. No.", "Name", "Designation", "Signature"]):
        format_cell(staff_table.rows[0].cells[i], h, bold=True, size=9)
        staff_table.rows[0].cells[i].width = staff_widths[i]

    for idx, member in enumerate(profile["staff"], start=1):
        row_cells = staff_table.add_row().cells
        # Keep Signature field explicitly blank ("")
        vals = [str(idx), member["name"], member["designation"], ""]
        for i, val in enumerate(vals):
            format_cell(row_cells[i], val, size=9)
            row_cells[i].width = staff_widths[i]

    # CRC Attendance Section
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run("CASE REVIEW COMMITTEE ATTENDANCE")
    r.font.name = 'Calibri'
    r.font.size = Pt(10)
    r.font.bold = True

    crc_table = doc.add_table(rows=1, cols=4)
    crc_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    crc_table.style = 'Table Grid'
    
    for i, h in enumerate(["S. No.", "Name", "Designation", "Signature*"]):
        format_cell(crc_table.rows[0].cells[i], h, bold=True, size=9)
        crc_table.rows[0].cells[i].width = staff_widths[i]

    for idx, member in enumerate(profile["crc"], start=1):
        row_cells = crc_table.add_row().cells
        # Keep Signature field explicitly blank ("")
        vals = [str(idx), member["name"], member["designation"], ""]
        for i, val in enumerate(vals):
            format_cell(row_cells[i], val, size=9)
            row_cells[i].width = staff_widths[i]

    # Mandatory Signature Note & Focal Person Table
    p_note = doc.add_paragraph()
    p_note.paragraph_format.space_before = Pt(4)
    p_note.paragraph_format.space_after = Pt(4)
    r_note = p_note.add_run("*Signatures are not mandatory for online participants")
    r_note.font.name = 'Calibri'
    r_note.font.size = Pt(8.5)
    r_note.font.italic = True

    focal_table = doc.add_table(rows=2, cols=3)
    focal_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    focal_table.style = 'Table Grid'
    focal_widths = [Inches(2.8), Inches(2.1), Inches(2.1)]
    
    format_cell(focal_table.rows[0].cells[0], "SJDA Focal Person", bold=True, size=9)
    format_cell(focal_table.rows[0].cells[1], "Signature", bold=True, size=9)
    format_cell(focal_table.rows[0].cells[2], "Date", bold=True, size=9)

    format_cell(focal_table.rows[1].cells[0], "Assistant Program Manager", size=9)
    # Keep Focal Person Signature explicitly blank ("")
    format_cell(focal_table.rows[1].cells[1], "", size=9)
    format_cell(focal_table.rows[1].cells[2], meeting_date.strftime("%d/%m/%Y"), size=9)

    for r_idx in range(2):
        for c_idx in range(3):
            focal_table.rows[r_idx].cells[c_idx].width = focal_widths[c_idx]

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- Streamlit UI Form ---
st.title("SJDA CRC Meeting Minutes Generator")

selected_location = st.selectbox("Select Area / Council Location", list(AREA_PROFILES.keys()))
profile = AREA_PROFILES[selected_location]

st.subheader("1. Meeting Details")
col1, col2, col3 = st.columns(3)
with col1:
    meeting_date = st.date_input("Meeting Date")
with col2:
    meeting_time = st.text_input("Time", value="04:00 PM to 05:00 PM")
with col3:
    venue = st.text_input("Venue / Zoom Link", value=profile["venue"])

st.subheader("2. Add FDP Beneficiaries")
num_fdps = st.number_input("How many FDP decisions to enter?", min_value=1, max_value=20, value=2)

fdp_data = []
for i in range(int(num_fdps)):
    st.markdown(f"**Beneficiary #{i+1}**")
    c1, c2, c3, c4, c5, c6, c7 = st.columns([2, 1.2, 1.2, 1.2, 1.2, 1.2, 1.2])
    
    name = c1.text_input(f"Name #{i+1}", key=f"name_{i}")
    fam_id = c2.text_input(f"Family ID #{i+1}", value="PE-", key=f"id_{i}")
    pov = c3.selectbox(f"Poverty Level #{i+1}", ["Level -1", "Level -2", "Level -3", "Level -4"], key=f"pov_{i}")
    support = c4.number_input(f"Total Support #{i+1}", value=75000, step=1000, key=f"sup_{i}")
    loan = c5.number_input(f"Loan #{i+1}", value=0, step=1000, key=f"loan_{i}")
    grant = c6.number_input(f"Grant #{i+1}", value=75000, step=1000, key=f"grant_{i}")
    status = c7.selectbox(f"Status #{i+1}", ["", "Approved", "Rejected", "Deferred"], key=f"stat_{i}")
    
    fdp_data.append({
        "name": name,
        "family_id": fam_id,
        "poverty_level": pov,
        "support": support,
        "loan": loan,
        "grant": grant,
        "status": status
    })

st.markdown("---")
if st.button("Generate Word Document", type="primary"):
    docx_file = generate_docx(selected_location, meeting_date, meeting_time, venue, fdp_data, profile)
    st.download_button(
        label="Download Meeting Minutes (.docx)",
        data=docx_file,
        file_name=f"{selected_location}_CRC_Minutes_{meeting_date}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )