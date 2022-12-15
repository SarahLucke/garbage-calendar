import requests

from openpyxl.styles.fills import PatternFill
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from datetime import datetime
from datetime import date

import calendar
from calendar import monthrange


class TableCreator:
    weekdays = ["Mo", "Di", "Mi", "Do", "Fr", "Sa", "So"]
    type_dict = {
        "INGOL_REST": {
            "cellValue": "R",
            "cellColor": "808080",
            "description": "Restmüll",
        },
        "INGOL_BIO": {
            "cellValue": "B",
            "cellColor": "33cc00",
            "description": "Biomüll",
        },
        "INGOL_GELB": {
            "cellValue": "G",
            "cellColor": "FFDD00",
            "description": "gelber Sack",
        },
        "INGOL_PAP": {
            "cellValue": "P",
            "cellColor": "00CCFF",
            "description": "Papiermüll",
        },
    }

    def __init__(self, user_input):
        self.thisYear = date.today().year
        # use base_urls.get_url_monthly_dates(year, month) instead
        params = {
            "idx": "termins",
            "city_id": user_input.city_id,
            "area_id": user_input.area_id,
            "ws": "3",
        }
        # get calendar info:
        response = requests.get("https://ingol.jumomind.com/webservice.php", params)
        self.data = response.json()[0]["_data"]

    def create_excel(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "calendar"

        yOffset = 2
        xOffset = 1
        rowspan = 2
        # write heading:
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=15)
        cell = ws.cell(row=1, column=1)
        cell.value = "Müllabfuhr " + str(self.thisYear)
        cell.alignment = Alignment(horizontal="left")
        cell.font = Font(bold=True, size=32)
        ws.row_dimensions[1].height = 40

        # border definition:
        border = Side(border_style="hair", color="000000")
        full_border = Border(top=border, left=border, right=border, bottom=border)
        top_bottom_left_border = Border(top=border, left=border, bottom=border)
        top_bottom_right_border = Border(top=border, right=border, bottom=border)
        top_bottom_border = Border(top=border, bottom=border)
        top_border = Border(top=border)

        # write numbers at border:
        ws.column_dimensions[get_column_letter(xOffset)].width = 4
        # first empty cell:
        cell = ws.cell(row=yOffset, column=xOffset)
        cell.border = full_border
        # number cells:
        for num in range(1, 32):
            cell = ws.cell(row=num + yOffset, column=xOffset)
            cell.value = num
            cell.alignment = Alignment(horizontal="center")
            cell.font = Font(bold=True)
            cell.border = full_border

        # for loop over dict/array
        # write legend:
        row = 33 + yOffset
        col = xOffset + 1
        for garbage_type in self.type_dict:
            cellInfo = self.type_dict.get(garbage_type)
            cell = ws.cell(row=row, column=col)
            cell.alignment = Alignment(horizontal="center")
            cell.value = cellInfo.get("cellValue")
            colorFill = PatternFill(
                fgColor=cellInfo.get("cellColor"), fill_type="solid"
            )
            cell.fill = colorFill
            cell = ws.cell(row=row, column=col + 1)
            cell.value = ": " + cellInfo.get("description")
            col = col + rowspan + 1

        counter = 0

        for month in range(1, 13):
            # write month names on top:
            xPos = month * rowspan + month - rowspan + xOffset
            cell = ws.cell(row=yOffset, column=xPos)
            cell.value = calendar.month_name[month]
            cell.alignment = Alignment(horizontal="center")
            cell.font = Font(bold=True)
            cell.border = full_border
            ws.merge_cells(
                start_row=yOffset,
                start_column=xPos,
                end_row=yOffset,
                end_column=xPos + rowspan,
            )
            maxDays = monthrange(self.thisYear, month)
            # format weekday column:
            ws.column_dimensions[get_column_letter(xPos)].width = 4
            for colCounter in range(1, rowspan + 1):
                ws.column_dimensions[get_column_letter(xPos + colCounter)].width = 6

            # loop through days of month:
            for day in range(1, maxDays[1] + 1):
                runningDate = date(self.thisYear, month, day)
                yPos = yOffset + day
                # write weekday:
                cell = ws.cell(row=yPos, column=xPos)
                cell.value = self.weekdays[runningDate.weekday()]
                # print "Sa" and "So" bold:
                if runningDate.weekday() == 5 or runningDate.weekday() == 6:
                    cell.font = Font(bold=True)
                # put borders
                cell.border = full_border
                for colCounter in range(1, rowspan + 1):
                    cell = ws.cell(row=yPos, column=xPos + colCounter)
                    if colCounter == 1:
                        cell.border = top_bottom_left_border
                    elif colCounter == rowspan:
                        cell.border = top_bottom_right_border
                    else:
                        cell.border = top_bottom_border

                # skip (continue counting) if wrong year:
                while (
                    datetime.strptime(self.data[counter]["cal_date"], "%Y-%m-%d")
                    .date()
                    .year
                    != self.thisYear
                ):
                    counter += 1
                    calendarDate = datetime.strptime(
                        self.data[counter]["cal_date"], "%Y-%m-%d"
                    ).date()

                # write all garbage info for the day:
                if (
                    datetime.strptime(self.data[counter]["cal_date"], "%Y-%m-%d").date()
                    == runningDate
                ):
                    xcounter = 1

                    while (
                        datetime.strptime(
                            self.data[counter]["cal_date"], "%Y-%m-%d"
                        ).date()
                        == runningDate
                    ):
                        garbage_type = self.data[counter]["cal_garbage_type"]

                        # format cell:
                        cell = ws.cell(row=yPos, column=xPos + xcounter)
                        cell.alignment = Alignment(horizontal="center")
                        cellInfo = self.type_dict.get(garbage_type)
                        cell.value = cellInfo.get("cellValue")
                        colorFill = PatternFill(
                            fgColor=cellInfo.get("cellColor"), fill_type="solid"
                        )
                        cell.fill = colorFill
                        counter += 1
                        xcounter += 1

                    # merge cells if only one entry:
                    if xcounter == xOffset + 1:
                        ws.merge_cells(
                            start_row=yPos,
                            start_column=xPos + 1,
                            end_row=yPos,
                            end_column=xPos + 2,
                        )

            # put border in last cell of month:
            for colCounter in range(0, rowspan + 1):
                cell = ws.cell(row=yOffset + 32, column=xPos + colCounter)
                cell.border = top_border

        # define print area
        ws.print_area = "A1:AK35"
        # set to page landscape mode
        ws.set_printer_settings(
            paper_size=openpyxl.worksheet.worksheet.Worksheet.PAPERSIZE_A4,
            orientation=openpyxl.worksheet.worksheet.Worksheet.ORIENTATION_LANDSCAPE,
        )
        # define scaling when printing
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToHeight = False

        wb.save("garbageCalendar" + str(self.thisYear) + ".xls")
