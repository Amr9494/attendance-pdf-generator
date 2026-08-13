import io
import calendar
import datetime
import streamlit as st
from openpyxl import load_workbook

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, PageBreak

# Page setup
st.set_page_config(page_title="Attendance PDF Generator", layout="centered")

st.title("📋 Attendance Sheet Generator")
st.write("Upload your student Excel file to generate printable PDF attendance sheets.")

# Sidebar Configuration Controls
st.sidebar.header("Layout Settings")
max_students = st.sidebar.slider("Max Students per Page", min_value=5, max_value=20, value=12)
student_col_width = st.sidebar.slider("Student Name Column Width", min_value=80, max_value=200, value=120)

uploaded_file = st.file_uploader("Choose an Excel file (.xlsx)", type=["xlsx"])

def generate_pdf(file_bytes, max_students_per_page, student_name_width):
    # Load Excel from upload stream
    wb = load_workbook(filename=io.BytesIO(file_bytes), data_only=True)
    sheet = wb.active
    students = []
    for row in sheet.iter_rows(min_row=2, min_col=1, max_col=1, values_only=True):
        if row[0] is not None and str(row[0]).strip() != "":
            students.append(str(row[0]).strip())

    if not students:
        return None

    pdf_buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        pdf_buffer,
        pagesize=A4,
        leftMargin=28,
        rightMargin=28,
        topMargin=20,
        bottomMargin=20
    )

    story = []
    start_date = datetime.date(2026, 8, 9)
    months = [(2026, 8), (2026, 9), (2026, 10), (2026, 11)]

    student_chunks = [
        students[i:i + max_students_per_page] 
        for i in range(0, len(students), max_students_per_page)
    ]

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'MonthTitle', parent=styles['Heading1'], fontName='Helvetica-Bold',
        fontSize=24, leading=28, textColor=colors.HexColor('#000000'),
        alignment=1, spaceAfter=10
    )
    student_style = ParagraphStyle(
        'StudentName', parent=styles['Normal'], fontName='Helvetica-Bold',
        fontSize=11, leading=13, textColor=colors.HexColor('#000000')
    )
    header_date_style = ParagraphStyle(
        'HeaderDate', parent=styles['Normal'], fontName='Helvetica-Bold',
        fontSize=16, leading=18, textColor=colors.HexColor('#000000'), alignment=1
    )

    printable_width = 538
    printable_height = 700
    header_row_height = 35
    student_row_height = (printable_height - header_row_height) / max_students_per_page

    for year, month in months:
        month_name = calendar.month_name[month]
        sundays = []
        cal = calendar.monthcalendar(year, month)
        for week in cal:
            day = week[6]
            if day != 0:
                d = datetime.date(year, month, day)
                if d >= start_date:
                    sundays.append(day)

        num_sundays = len(sundays)

        for chunk in student_chunks:
            story.append(Paragraph(month_name, title_style))
            header_row = [Paragraph("", header_date_style)] + [
                Paragraph(str(day), header_date_style) for day in sundays
            ]
            table_data = [header_row]

            for student_name in chunk:
                table_data.append([Paragraph(student_name, student_style)] + [""] * num_sundays)

            for _ in range(max_students_per_page - len(chunk)):
                table_data.append([""] + [""] * num_sundays)

            remaining_width = printable_width - student_name_width
            day_col_width = remaining_width / num_sundays if num_sundays > 0 else remaining_width
            col_widths = [student_name_width] + [day_col_width] * num_sundays
            row_heights = [header_row_height] + [student_row_height] * max_students_per_page

            t = Table(table_data, colWidths=col_widths, rowHeights=row_heights)
            t.setStyle(TableStyle([
                ('GRID', (0, 0), (-1, -1), 3.5, colors.HexColor('#000000')),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('ALIGN', (0, 0), (0, -1), 'LEFT'),
                ('ALIGN', (1, 0), (-1, -1), 'CENTER'),
                ('BOTTOMPADDING', (0, 1), (0, -1), 6),
                ('LEFTPADDING', (0, 1), (0, -1), 8),
            ]))

            story.append(t)
            story.append(PageBreak())

    if story and isinstance(story[-1], PageBreak):
        story.pop()

    doc.build(story)
    pdf_buffer.seek(0)
    return pdf_buffer

if uploaded_file is not None:
    if st.button("Generate PDF Attendance Sheet"):
        with st.spinner("Generating PDF..."):
            pdf_data = generate_pdf(uploaded_file.getvalue(), max_students, student_col_width)
            if pdf_data:
                st.success("PDF generated successfully!")
                st.download_button(
                    label="📥 Download PDF",
                    data=pdf_data,
                    file_name="Attendance_Sheets.pdf",
                    mime="application/pdf"
                )
            else:
                st.error("No student names found in Column A of the uploaded file.")