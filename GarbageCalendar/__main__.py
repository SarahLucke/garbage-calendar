#!/usr/bin/python
import sys

from datetime import date

from .gui import Gui
from .user_input import UserInput
from .table_creator import TableCreator

if __name__ == "__main__":
    # if no arguments provided, open GUI:
    if len(sys.argv) == 1:
        user_input = Gui()
    else:
        user_input = UserInput(
            city=sys.argv[1],
            street=sys.argv[2],
            house_number=sys.argv[3],
            year=int(sys.argv[4]) if len(sys.argv) > 4 else date.today().year,
        )
    if user_input.location is None or user_input.location.area_id.__eq__(""):
        sys.exit("the data you entered is not complete or faulty")

    table_creator = TableCreator(user_input)
    table_creator.create_pdf()
    table_creator.create_excel()

    print("finished")
