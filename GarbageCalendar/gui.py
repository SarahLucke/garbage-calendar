import tkinter as tk
from tkinter import ttk

from datetime import date

from .location import Location


class Gui:
    city = ""
    street = ""
    house_number = ""
    year = 0

    def __init__(self):
        ## open gui (pipe response from api through) and ask for street and housenumber.
        ## interact with user:
        self.query_window = tk.Tk()
        self.query_window.title("Additional information required")

        tk.Label(self.query_window, text="Year: ", anchor="e").grid(row=0, column=1)
        tk.Label(self.query_window, text="City: ", anchor="e").grid(row=1, column=1)
        tk.Label(self.query_window, text="Street: ", anchor="e").grid(row=2, column=1)
        tk.Label(self.query_window, text="House Number: ", anchor="e").grid(
            row=3, column=1
        )

        current_year=date.today().year
        self.year_input = tk.Spinbox(self.query_window, width=27, from_=current_year, to=current_year+1, textvariable=current_year, state="readonly")
        self.year_input.grid(row=0, column=2)
        self.city_chooser = ttk.Combobox(
            self.query_window, width=27, state="readonly"
        )
        self.city_chooser["values"] = tuple(Location.endpoints.keys())
        self.city_chooser.grid(row=1, column=2)
        self.city_chooser.bind("<<ComboboxSelected>>", self.setCity)

        self.street_chooser = ttk.Combobox(
            self.query_window, width=27, textvariable=tk.StringVar()
        )
        self.street_chooser.grid(row=2, column=2)

        self.number_chooser = ttk.Combobox(
            self.query_window, width=27, state="readonly"
        )
        self.number_chooser.grid(row=3, column=2)

        self.street_chooser.bind("<<ComboboxSelected>>", self.setHouseNumbers)
        self.street_chooser.bind("<KeyRelease>", self.handle_street_keyrelease)

        self.query_window.mainloop()
    
    def setCity(self, event):
        self.city = self.city_chooser.get()
        self.location = Location(self.city)
        self.street_chooser["values"] = tuple(self.location.get_street_names_like(None))

    def setHouseNumbers(self, event):
        self.street = self.street_chooser.get()
        self.number_chooser["values"] = tuple(
            self.location.get_street_house_numbers(self.street)
        )
        self.number_chooser.bind("<<ComboboxSelected>>", self.startGeneration)

    def handle_street_keyrelease(self, event):
        typed = self.street_chooser.get()
        self.street_chooser["values"] = tuple(
            self.location.get_street_names_like(typed)
        )
        self.street_chooser.event_generate("<Down>")

    def startGeneration(self, event):
        self.year = int(self.year_input.get())
        self.house_number = self.number_chooser.get()
        self.area_id = self.location.get_ref_area(self.street, self.house_number)
        self.query_window.destroy()
