# import libraries
import pandas as pd
from openpyxl import load_workbook
from reportlab.lib.pagesizes import letter, landscape
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader

# define function
def excel_to_pdf(excel_file, pdf_file, start_row, end_row, start_col, end_col):
    workbook = load_workbook(excel_file)
    sheet = workbook.active

    data_frame = sheet.values
    columns = next(data_frame)[start_col:end_col + 1]
    rows = list(data_frame)[start_row:end_row + 1]
    df = pd.DataFrame(rows, columns=columns)

    c = canvas.Canvas(pdf_file, pagesize=landscape(letter))  # landscape or potrait
    width, height = landscape(letter)

    c.setFont("Times New Roman", 12) # select font and size

    x_offset = 40  # units are point(1 point = 1/72 inch)
    y_offset = height - 40
    row_height = 20
    col_width = width / (end_col - start_col + 1)

    for i, col in enumerate(df.columns):
        c.drawString(x_offset + i * col_width, y_offset, col)

    for row in df.itertuples():
        y_offset -= row_height
        for i, value in enumerate(row[1:]):
            cell_value = str(value)
            if cell_value.startswith("http") or cell_value.startswith("www"):
                c.drawString(x_offset + i * col_width, y_offset, "Check following link")
                link = cell_value if cell_value.startswith("http") else f"http://{cell_value}"
                c.linkURL(link,
                          (x_offset + i * col_width, y_offset, x_offset + (i + 1) * col_width, y_offset + row_height))
            else:
                c.drawString(x_offset + i * col_width, y_offset, cell_value)


    c.save()


excel_to_pdf("file_name.xlsx", "print.pdf", 0, 10, 0, 5)
