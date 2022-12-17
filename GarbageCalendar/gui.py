import tkinter as tk
from tkinter import ttk  # additional tkinter elements like combo boxes

from .location import Location


class Gui:
    street = ""

    def __init__(self):
        self.city = "Ingolstadt"
        ## open gui (pipe response from api through) and ask for street and housenumber.
        self.location = Location(city)
        ## interact with user:
        self.query_window = tk.Tk()
        self.query_window.title("Additional information required")

        tk.Label(self.query_window, text="Street: ", anchor="e").grid(row=1, column=1)
        tk.Label(self.query_window, text="House Number: ", anchor="e").grid(
            row=2, column=1
        )

        self.street_chooser = ttk.Combobox(
            self.query_window, width=27, textvariable=tk.StringVar()
        )
        self.street_chooser["values"] = tuple(self.location.get_street_names_like(None))
        self.street_chooser.grid(row=1, column=2)

        self.number_chooser = ttk.Combobox(
            self.query_window, width=27, textvariable=tk.StringVar()
        )
        self.number_chooser.grid(row=2, column=2)

        self.street_chooser.bind("<<ComboboxSelected>>", self.setHouseNumbers)
        self.street_chooser.bind("<KeyRelease>", self.handle_street_keyrelease)

        self.query_window.mainloop()

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
        self.house_number = self.number_chooser.get()
        self.area_id = self.location.get_ref_area(self.street, self.house_number)
        self.query_window.destroy()
