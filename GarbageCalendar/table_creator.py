import requests

from openpyxl.styles.fills import PatternFill
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from datetime import datetime
from datetime import date

import calendar
from calendar import monthrange

import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
from matplotlib import colors

from .base_requests import BaseRequests


class TableCreator(BaseRequests):
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
        self.thisYear = user_input.year
        if self.thisYear < date.today().year:
            self.thisYear = date.today().year
        self.end_point = user_input.location.end_point
        self.city_id = user_input.location.city_id
        self.area_id = user_input.location.area_id

        self.title = "Müllabfuhr {}".format(str(self.thisYear))
        self.file_name = "garbageCalendar{}".format(str(self.thisYear))
        self.address = "{} {}\n{}".format(
            user_input.street, user_input.house_number, user_input.city
        )
        self.num_months = 12
        self.num_days = 31

    def create_pdf(self):
        empty_cell_color = "#ffffff"

        plt.figure(
            linewidth=1,
            tight_layout={"pad": 1.5},
        )
        ax = plt.gca()
        ax.axis("off")
        plt.box(on=None)  # Add title
        plt.suptitle(self.title)  # Add footer
        plt.figtext(
            0.95,
            0.05,
            self.address,
            horizontalalignment="right",
            size=6,
            weight="light",
        )

        # Add a table at the bottom of the axes
        calendar_tab = plt.table(
            cellText=[[""] * (3 * self.num_months)] * self.num_days,
            rowLabels=range(1, self.num_days + 1),
            rowLoc="center",
            colLabels=[""] * (3 * self.num_months),
            loc="upper center",
            colWidths=[0.03] * (3 * self.num_months),
        )
        calendar_tab.auto_set_font_size(False)
        calendar_tab.set_fontsize(6)

        ypos_header = 0
        for xpos in range(0, self.num_months):
            month = xpos + 1
            xpos = xpos * 3
            cell = calendar_tab[ypos_header, xpos + 2]
            cell.visible_edges = "B"
            cell = calendar_tab[ypos_header, xpos]
            cell.visible_edges = "B"
            cell = calendar_tab[ypos_header, xpos + 1]
            cell.visible_edges = "B"
            if ypos_header == ypos_header:
                month_name = calendar.month_name[month]
                calendar_tab[ypos_header, xpos + 2].set_text_props(ha="right")
                calendar_tab[ypos_header, xpos + 2].get_text().set_text(month_name)

        for xpos in range(0, self.num_months):
            month = xpos + 1
            xpos = xpos * 3

            maxDays = monthrange(self.thisYear, month)
            # get calendar data
            response = requests.get(self.get_url_monthly_dates(self.thisYear, month))
            data = response.json()["dates"]

            for day in range(1, self.num_days + 1):
                if day < maxDays[1] + 1:
                    runningDate = date(self.thisYear, month, day)
                    # write week day:
                    week_day = self.weekdays[runningDate.weekday()]
                    cell = calendar_tab[day, xpos]
                    cell.get_text().set_text(week_day)
                    cell.set_text_props(ha="left")
                    if runningDate.weekday() == 5 or runningDate.weekday() == 6:
                        cell.set_text_props(weight="bold")

                    if len(data) == 0:
                        continue

                    # write all garbage info for the day:
                    for date_info in data:
                        if (
                            datetime.strptime(date_info["day"], "%Y-%m-%d").date()
                            == runningDate
                        ):
                            garbage_type = date_info["trash_name"]

                            # format cell:
                            xcounter = 1
                            cell = calendar_tab[day, xpos + xcounter]
                            while (
                                xcounter < 2
                                and str(colors.to_hex(cell.get_facecolor()))
                                != empty_cell_color
                            ):
                                xcounter += 1
                                cell = calendar_tab[day, xpos + xcounter]
                            cellInfo = self.type_dict.get(garbage_type)
                            cell.get_text().set_text(cellInfo.get("cellValue"))
                            cell.set_text_props(ha="center")
                            cell.set_facecolor(f'#{cellInfo.get("cellColor")}')
                            cell.set_edgecolor("none")

                # format cells
                cell2 = calendar_tab[day, xpos + 2]
                cell2.visible_edges = "TBR"
                cell = calendar_tab[day, xpos + 1]
                # no color in cell1
                if str(colors.to_hex(cell.get_facecolor())) == empty_cell_color:
                    cell.visible_edges = "TB"
                else:
                    # color in cell1, but not in cell2 -> merge
                    if str(colors.to_hex(cell2.get_facecolor())) == empty_cell_color:
                        cell2.set_facecolor(str(colors.to_hex(cell.get_facecolor())))
                        cell2.set_edgecolor("none")
                        cell.set_text_props(ha="right")
                cell = calendar_tab[day, xpos]
                cell.visible_edges = "LTB"

        with PdfPages(self.file_name + ".pdf") as export_pdf:
            plt.draw()
            export_pdf.savefig()
            plt.close()

    def create_excel(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "calendar"

        empty_cell_color = "00000000"

        yOffset = 2
        xOffset = 1
        rowspan = 2
        # write heading:
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=15)
        cell = ws.cell(row=1, column=1)
        cell.value = self.title
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
        for num in range(1, self.num_days + 1):
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

        for month in range(1, self.num_months + 1):
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

            # get calendar data
            response = requests.get(self.get_url_monthly_dates(self.thisYear, month))
            data = response.json()["dates"]

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

                if len(data) == 0:
                    continue

                # write all garbage info for the day:
                for date_info in data:
                    if (
                        datetime.strptime(date_info["day"], "%Y-%m-%d").date()
                        == runningDate
                    ):
                        garbage_type = date_info["trash_name"]

                        # format cell:
                        xcounter = 1
                        cell = ws.cell(row=yPos, column=xPos + xcounter)
                        while (
                            xcounter < rowspan
                            and cell.fill.start_color.index != empty_cell_color
                        ):
                            xcounter += 1
                            cell = ws.cell(row=yPos, column=xPos + xcounter)
                        cell.alignment = Alignment(horizontal="center")
                        cellInfo = self.type_dict.get(garbage_type)
                        cell.value = cellInfo.get("cellValue")
                        colorFill = PatternFill(
                            # fgColor=date_info["color"],
                            fgColor=cellInfo.get("cellColor"),
                            fill_type="solid",
                        )
                        cell.fill = colorFill

                # merge cells if only one entry:
                cell = ws.cell(row=yPos, column=xPos + rowspan)
                if cell.fill.start_color.index == empty_cell_color:
                    ws.merge_cells(
                        start_row=yPos,
                        start_column=xPos + 1,
                        end_row=yPos,
                        end_column=xPos + rowspan,
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

        wb.save(self.file_name + ".xls")
