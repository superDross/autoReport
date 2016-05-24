import openpyxl
import sys
from openpyxl.styles import Border, Alignment, Font, Side
from openpyxl.styles.colors import BLACK

# alter formating if it is windows

def create_template(output_path=""):
    ''' Create a template report workbook
    '''
    # create worksheet
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = " "


    # merge cells
    ws.merge_cells('B1:J1')
    ws.merge_cells('F4:J4')
    ws.merge_cells('B16:E20')
    ws.merge_cells('B21:E21')
    ws.merge_cells('F16:J20')
    ws.merge_cells('B33:D33')
    ws.merge_cells('B47:J48')

    for num in range(4, 16):
        if num != 9 and num != 10:
            cell_range = "".join(("B", str(num), ":", "E", str(num)))
            ws.merge_cells(cell_range)   

    ws.merge_cells('B9:E10')

    for n in range(7, 16):
        cell_range = "".join(("F", str(n), ":", "J", str(n)))
        ws.merge_cells(cell_range)


    # style variables
    center = Alignment(horizontal="center")
    center_v = Alignment(vertical="center")
    left = Alignment(horizontal="left")
    font = Font(size=14, color=BLACK)
    small_font = Font(size=8, color=BLACK)
    thin_border = Border(left=Side(border_style='thin'), right=Side(border_style='thin'), 
                         top=Side(border_style='thin'), bottom=Side(border_style='thin'))


    # apply border and alignment
    for i in range(7,21):
        ws.cell(row=i, column=2).border = thin_border
        ws.cell(row=i, column=11).border = Border(left=Side(border_style='thin'))
        ws.cell(row=i, column=3).border = thin_border
        ws.cell(row=i, column=4).border = thin_border
        ws.cell(row=i, column=5).border = thin_border
        ws.cell(row=i, column=6).border = thin_border
        ws.cell(row=i, column=7).border = thin_border
        ws.cell(row=i, column=8).border = thin_border
        ws.cell(row=i, column=9).border = thin_border
        ws.cell(row=i, column=10).border = thin_border
        ws.cell(row=i, column=2).alignment = center_v

    for i in range(6,11):
        ws.cell(row=10, column=i).border = Border(top=Side(border_style=None))
        ws.cell(row=9, column=i).border = Border(bottom=Side(border_style=None))

    ws.cell(row=10,column=10).border = Border(right=Side(border_style="thin"))
    ws.cell(row=9,column=10).border = Border(right=Side(border_style="thin"))


    for i in range(6,11):
        ws.cell(row=4, column=i).border = thin_border




    # stylise entry cells and their title cells
    for i in range(7,17):
        cell = ws.cell(row=i,column=6)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.font = Font(size=10, color=BLACK)
        title_cell = ws.cell(row=i, column=2)
        title_cell.font = Font(size=14, color=BLACK)

    
    comment_cell = ws.cell(row=16, column=6)
    comment_cell.font = Font(size=6, color=BLACK)
    comment_cell.style.alignment.wrap_text = True
    ws["F4"].alignment = Alignment(horizontal="center", vertical="center")
    ws["F4"].font = Font(size=10, color=BLACK)

    # change column width
    ws.column_dimensions["E"].width = 7
    ws.row_dimensions[2].height = 1
    ws.row_dimensions[3].height = 7
    ws.row_dimensions[5].height = 5

    # write to cells
    Title = ws['B1']
    ws['B1'] = "Variant Confirmation Report"
    Title.font = Font(size=20,underline='single')
    Title.alignment = center
    ws.row_dimensions[1].height = 25

    ID = ws['B4']
    ws['B4'] = "Sample ID:"
    ID.font = font 
    ID.alignment = left
    ws.row_dimensions[4].height = 18

    Describe = ws['B6']
    ws['B6'] = "Description:"
    Describe.font = font
    Describe.alignment = left
    ws.row_dimensions[6].height = 18

    Fluidigm = ws['B21']
    ws['B21'] = "Validation:"
    Fluidigm.font = font
    Fluidigm.alignment = left
    ws.row_dimensions[21].height = 18

    Validation = ws['B34']
    ws['B34'] = "Fluidigm Image:"
    Validation.font = font
    Validation.alignment = left
    ws.row_dimensions[34].height = 18

    ws['B7'] = "Gene"
    ws['B8'] = "Exon"
    ws['B9'] = "Sequence Variant"
    ws['B11'] = "Variant Location (GRCh37)"
    # ws.row_dimensions[11].height = 30
    ws['B12'] = "Allele Balance"
    ws['B13'] = "Allele Depth (REF,ALT)"
    # ws.row_dimensions[13].height = 30
    ws['B14'] = "Allele Frequency (ESP,ExAC,dbSNP)"
    # ws.row_dimensions[14].height = 30
    ws['B15'] = "Variant Found"
    ws['B16'] = "Comment"



    Footer = ws['B47']
    ws['B47'] = ("The following information is for research purpose only. Any decisions made on the information should be made by an appropriate\nresponsible clinician who may require further confirmation within a clinical laboratory.")
    Footer.font = Font(size=8)
    Footer.alignment = center
    Footer.style.alignment.wrap_text = True


    wb.save(output_path+"test_template.xlsx")

