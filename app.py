import io
import calendar
import datetime
import streamlit as st
from openpyxl import load_workbook
import pypdfium2 as pdfium

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, PageBreak

st.set_page_config(page_title="Attendance PDF Generator", layout="centered")

st.title("📋 Attendance Sheet Generator")
st.write("Upload your student Excel file to generate printable PDF attendance sheets.")

# Sidebar Configuration Controls
st.sidebar.header("Layout Settings")
max_students = st.sidebar.number_input("Max Students per Page", min_value=5, max_value=25, value=12, step=1)
student_col_width = st.sidebar.number_input("Student Name Column Width (pt)", min_value=60, max_value=250, value=120, step=5)

uploaded_file = st.file_uploader("Choose an Excel file (.xlsx)", type=["xlsx"])

def parse_students(file_bytes):
    """Extracts student names from Column A of the uploaded Excel file."""
    wb = load_workbook(filename=io.BytesIO(file_bytes), data_only=True)
    sheet = wb.active
    students = []
    for row in sheet.iter_rows(min_row=2, min_col=1, max_col=1, values_only=True):
        if row[0] is not None and str(row[0]).strip() != "":
            students.append(str(row[0]).strip())
    return students

def generate_pdf(students, max_students_per_page, student_name_width, preview_only=False):
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
    # If previewing, only generate August; otherwise generate August through November
    months = [(2026, 8)] if preview_only else [(2026, 8), (2026, 9), (2026, 10), (2026, 11)]

    student_chunks = [
        students[i:i + max_students_per_page] 
        for i in range(0, len(students), max_students_per_page)
    ]

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'MonthTitle', parent=styles['Heading1'], fontName='Helvetica-Bold',
        fontSize=22, leading=26, textColor=colors.HexColor('#000000'),
        alignment=1, spaceAfter=8
    )
    
    student_style = ParagraphStyle(
        'StudentName', parent=styles['Normal'], fontName='Helvetica-Bold',
        fontSize=10, leading=12, textColor=colors.HexColor('#000000'),
        wordWrap='CJK'  # Text wrapping for long names
    )
    
    header_date_style = ParagraphStyle(
        'HeaderDate', parent=styles['Normal'], fontName='Helvetica-Bold',
        fontSize=15, leading=17, textColor=colors.HexColor('#000000'), alignment=1
    )

    printable_width = 538
    printable_height = 700
    header_row_height = 32
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
                ('GRID', (0, 0), (-1, -1), 3.0, colors.HexColor('#000000')),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('ALIGN', (0, 0), (0, -1), 'LEFT'),
                ('ALIGN', (1, 0), (-1, -1), 'CENTER'),
                ('TOPPADDING', (0, 0), (-1, -1), 2),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
                ('LEFTPADDING', (0, 0), (-1, -1), 6),
                ('RIGHTPADDING', (0, 0), (-1, -1), 4),
            ]))

            story.append(t)
            
            if preview_only:
                break
            story.append(PageBreak())

    if story and isinstance(story[-1], PageBreak):
        story.pop()

    doc.build(story)
    pdf_buffer.seek(0)
    return pdf_buffer

def render_preview_image(pdf_bytes):
    """Converts the first page of the PDF buffer into an image for UI preview."""
    pdf = pdfium.PdfDocument(pdf_bytes)
    page = pdf[0]
    image = page.render(scale=2).to_pil()
    img_byte_arr = io.BytesIO()
    image.save(img_byte_arr, format='PNG')
    return img_byte_arr.getvalue()

if uploaded_file is not None:
    students = parse_students(uploaded_file.getvalue())
    if students:
        st.subheader("🖼️ Page 1 Layout Preview")
        st.info("Adjust the settings in the sidebar to dynamically re-size the table cells.")
        
        # Render dynamic image preview with updated parameter name
        preview_pdf = generate_pdf(students, max_students, student_col_width, preview_only=True)
        preview_img = render_preview_image(preview_pdf)
        st.image(preview_img, caption="Live Preview (First Page)", use_container_width=True)
        
        st.markdown("---")
        if st.button("🚀 Generate Full PDF (All Months)"):
            with st.spinner("Compiling full PDF..."):
                full_pdf = generate_pdf(students, max_students, student_col_width, preview_only=False)
                st.success("PDF generated successfully!")
                st.download_button(
                    label="📥 Download Complete PDF",
                    data=full_pdf,
                    file_name="Attendance_Sheets.pdf",
                    mime="application/pdf"
                )
    else:
        st.error("No student names were found in Column A of the uploaded file.")