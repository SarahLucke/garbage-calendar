import tkinter as tk
from tkinter import ttk

from datetime import date

from .location import Location


class Gui:
    city = ""
    street = ""
    house_number = ""
    year = 0
    location=None

    def __init__(self):
        ## open gui (pipe response from api through) and ask for street and housenumber.
        ## interact with user:
        self._query_window = tk.Tk()
        self._query_window.title("Additional information required")

        tk.Label(self._query_window, text="Year: ", anchor="e").grid(row=0, column=1)
        tk.Label(self._query_window, text="City: ", anchor="e").grid(row=1, column=1)
        tk.Label(self._query_window, text="Street: ", anchor="e").grid(row=2, column=1)
        tk.Label(self._query_window, text="House Number: ", anchor="e").grid(
            row=3, column=1
        )

        current_year = date.today().year
        self._year_input = tk.Spinbox(
            self._query_window,
            width=27,
            from_=current_year,
            to=current_year + 1,
            textvariable=current_year,
            state="readonly",
        )
        self._year_input.grid(row=0, column=2)
        self._city_chooser = ttk.Combobox(self._query_window, width=27, state="readonly")
        self._city_chooser["values"] = tuple(Location.endpoints.keys())
        self._city_chooser.grid(row=1, column=2)
        self._city_chooser.bind("<<ComboboxSelected>>", self.set_city)

        self._street_chooser = ttk.Combobox(
            self._query_window, width=27, textvariable=tk.StringVar()
        )
        self._street_chooser.grid(row=2, column=2)

        self._number_chooser = ttk.Combobox(
            self._query_window, width=27, state="readonly"
        )
        self._number_chooser.grid(row=3, column=2)

        self._street_chooser.bind("<<ComboboxSelected>>", self.set_house_numbers)
        self._street_chooser.bind("<KeyRelease>", self.handle_street_keyrelease)

        self._query_window.mainloop()

    def set_city(self, _event):
        self.city = self._city_chooser.get()
        self.location = Location(self.city)
        self._street_chooser["values"] = tuple(self.location.get_street_names_like(None))

    def set_house_numbers(self, _event):
        self.street = self._street_chooser.get()
        self._number_chooser["values"] = tuple(
            self.location.get_street_house_numbers(self.street)
        )
        self._number_chooser.bind("<<ComboboxSelected>>", self.start_generation)

    def handle_street_keyrelease(self, _event):
        typed = self._street_chooser.get()
        self._street_chooser["values"] = tuple(
            self.location.get_street_names_like(typed)
        )
        self._street_chooser.event_generate("<Down>")

    def start_generation(self, _event):
        self.year = int(self._year_input.get())
        self.house_number = self._number_chooser.get()
        self.area_id = self.location.get_ref_area(self.street, self.house_number)
        self._query_window.destroy()
