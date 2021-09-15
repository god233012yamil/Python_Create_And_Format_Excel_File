# XlsxWriter is a Python module that can be used to write text, numbers, formulas and
# hyperlinks to multiple worksheets in an Excel 2007+ XLSX file.
import xlsxwriter
from datetime import date
import os
from os import path
import sys

# Create the list of headers for the BOM.
list_of_headers = ['Seq.', 'Designator', 'Description', 'MFG P/N', 'Qty', 'Rev', 'MFG Name',
                   'Price Break', 'Qty Available at Digikey', 'Lead Time', 'Product Status', 'Operational Temp',
                   'Lead Status', 'ROHS Status', 'Datasheet', 'Web Page']


#
def create_excel(file_name):
    #
    try:
        # Create a workbook with the name to be 'file_name.xlsx'
        workbook = xlsxwriter.Workbook(file_name + '.xlsx')
        # Create a worksheet with the name to be 'file_name'
        worksheet = workbook.add_worksheet(file_name)
        # Freeze pane on the top row.
        # worksheet.freeze_panes(1, 0)
        # Freeze pane on the top row and left column.
        # worksheet.freeze_panes(1, 1)
        # Set the page orientation as landscape.
        worksheet.set_landscape()
        # Set the page view mode.
        worksheet.set_page_view()
        # Set the paper format to Tabloid (11 x 17).
        worksheet.set_paper(3)
        # Set the worksheet margins for the printed page.
        worksheet.set_margins(left=0.75, right=0.75, top=1.6, bottom=1.2)
        # Header definition
        header = "&C&\"Arial,Bold\"&20 BILL OF MATERIAL " \
                 "&L&\"Arial,Normal\"&18 \nP/N: {0} &R&\"Arial,Normal\"&18 \nBOM Rev: {1} " \
                 "&L&\"Arial,Normal\"&18 \nDescription: {2} &R&\"Arial,Normal\"&18 \nPCB Rev: {3} " \
                 "&R&\"Arial,Normal\"&18 \nFW Rev: {4} ".format("TBD", "TBD", "TBD", "TBD", "TBD")
        # Set the header.
        worksheet.set_header(header)
        # Footer definition
        footer = "&L&\"Arial,Normal\"&14BOM Doc. Number/File Name: {0} &R&\"Arial,Normal\"&14&P of &N " \
                 "&L&\"Arial,Normal\"&14 \nBOM Doc. FW Doc. Number/File Name: {1} " \
                 "&L&\"Arial,Normal\"&14 \nDate: {2} ".format("TBD", "TBD", date.today())
        # Set the printed page footer caption and options.
        worksheet.set_footer(footer)
        # Set the number of rows to repeat at the top of each printed page.
        worksheet.repeat_rows(0)
        # Set row height.
        worksheet.set_default_row(30)

        # Create a cell format for the headers.
        header_cell_format = workbook.add_format({'bold': True,
                                                  'font_color': 'white',
                                                  'font_name': 'Arial',
                                                  'font_size': 12,
                                                  'align': 'center',
                                                  'valign': 'vcenter',
                                                  'border': 1,
                                                  'text_wrap': True,
                                                  'bg_color': '#3399ff'})
        # Create a cell format for the headers.
        header_cell_format_left = workbook.add_format({'bold': True,
                                                       'font_color': 'white',
                                                       'font_name': 'Arial',
                                                       'font_size': 12,
                                                       'align': 'left',
                                                       'valign': 'vcenter',
                                                       'border': 1,
                                                       'text_wrap': True,
                                                       'bg_color': '#3399ff'})
        # Create a cell format for regular cells.
        cell_format_align_center = workbook.add_format({'bold': False,
                                                        'font_color': 'black',
                                                        'font_name': 'Arial',
                                                        'font_size': 10,
                                                        'align': 'center',
                                                        'valign': 'vcenter',
                                                        'border': 1,
                                                        'text_wrap': True})
        # Create a cell format for regular cells.
        cell_format_align_left = workbook.add_format({'bold': False,
                                                      'font_color': 'black',
                                                      'font_name': 'Arial',
                                                      'font_size': 10,
                                                      'align': 'left',
                                                      'valign': 'vcenter',
                                                      'border': 1,
                                                      'text_wrap': True})
        # Create a cell format for cells with link.
        cell_format_url = workbook.add_format({'bold': False,
                                               'font_color': 'blue',
                                               'font_name': 'Arial',
                                               'font_size': 10,
                                               'align': 'center',
                                               'valign': 'vcenter',
                                               'border': 1,
                                               'text_wrap': True,
                                               'underline': True})

        # Iterate over each column.
        for col in range(len(list_of_headers)):
            # Add the headers with the appropriate cell format.
            worksheet.write(0, col, list_of_headers[col], header_cell_format)

            # If the column header or name is:
            if list_of_headers[col] == "Seq.":
                # Set column width for mfg part number column
                worksheet.set_column(col, col, 6)

            # If the column header or name is:
            if list_of_headers[col] == "Designator":
                # Add headers to BOM with an appropriate cell format.
                worksheet.write(0, col, list_of_headers[col], header_cell_format_left)
                # Set column width for mfg part number column
                worksheet.set_column(col, col, 50)

            # If the column header or name is:
            if list_of_headers[col] == "Description":
                # Add headers to BOM with an appropriate cell format.
                worksheet.write(0, col, list_of_headers[col], header_cell_format_left)
                # Set column width for description column
                worksheet.set_column(col, col, 50)

            # If the column header or name is:
            if list_of_headers[col] == "MFG P/N" or list_of_headers[col] == "Manufacture P/N":
                # Add headers to BOM with an appropriate cell format.
                worksheet.write(0, col, list_of_headers[col], header_cell_format_left)
                # Set column width for mfg part number column
                worksheet.set_column(col, col, 28)

            # If the column header or name is:
            if list_of_headers[col] == "Qty":
                # Set column width for mfg part number column
                worksheet.set_column(col, col, 6)

            # If the column header or name is:
            if list_of_headers[col] == "Rev":
                # Set column width for Rev column
                worksheet.set_column(col, col, 6)

            # If the column header or name is:
            if list_of_headers[col] == "MFG Name":
                # Set column width for mfg name column
                worksheet.set_column(col, col, 15, options={'hidden': False})

            # If the column header or name is:
            if list_of_headers[col] == "Qty Available at Digikey":
                # Set column width for mfg part number column
                worksheet.set_column(col, col, 11, options={'hidden': True})

            # If the column header or name is:
            if list_of_headers[col] == "Lead Time":
                # Set column width for mfg part number column
                worksheet.set_column(col, col, 15, options={'hidden': False})

            # If the column header or name is:
            if list_of_headers[col] == "Product Status":
                # Set column width for Product Status column
                worksheet.set_column(col, col, 10, options={'hidden': True})

            # If the column header or name is:
            if list_of_headers[col] == "Operational Temp":
                # Set column width for Operational Temp column
                worksheet.set_column(col, col, 14, options={'hidden': True})

            # If the column header or name is:
            if list_of_headers[col] == "Price Break":
                # Set column width for Price Break column
                worksheet.set_column(col, col, 14, options={'hidden': True})

            # If the column header or name is:
            if list_of_headers[col] == "Datasheet":
                # Set column width for datasheet column.
                worksheet.set_column(col, col, 12, options={'hidden': True})

            # If the column header or name is:
            if list_of_headers[col] == "Web Page":
                # Set column width for URL column.
                worksheet.set_column(col, col, 8, options={'hidden': True})

            # If the column header or name is:
            if list_of_headers[col] == "Lead Status":
                # Set column width for Lead Status column
                worksheet.set_column(col, col, 15, options={'hidden': False})

            # If the column header or name is:
            if list_of_headers[col] == "ROHS Status":
                # Set column width for ROHS Status column
                worksheet.set_column(col, col, 12, options={'hidden': True})

        # Close the workbook.
        workbook.close()

        # Open/Launch the created excel file.
        if path.exists(file_name + ".xlsx"):
            os.startfile(file_name + ".xlsx")
        else:
            print("Error opening file {0}".format(file_name + ".xlsx"))
            sys.exit(1)

    except xlsxwriter.exceptions.FileCreateError:
        print("Error: During file creation.")


def main():
    create_excel("test")


if __name__ == '__main__':
    main()
