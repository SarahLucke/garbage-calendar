import requests

from typing import Optional

from .base_requests import BaseRequests


class Location(BaseRequests):
    city_id = ""

    def __init__(self, city: str):
        super().__init__(city)
        response = requests.get(self.get_url_streets())
        self.data = response.json()

    def get_street_house_numbers(self, street_name: str):
        result = []
        for street in self.data:
            if street["name"] == street_name:
                for house_number in street["houseNumbers"]:
                    result.append(house_number[0])
                break
        return result

    def get_street_names_like(self, typed: Optional[str]):
        result = []
        if typed is None or typed == "":
            return [street["name"] for street in self.data]
        for street in self.data:
            if street["name"].startswith(typed):
                result.append(street["name"])
        return result

    def get_ref_area(self, selected_street_name: str, selected_house_number: str):
        for street in self.data:
            if street["name"] == selected_street_name:
                for number in street["houseNumbers"]:
                    if number[0].__eq__(selected_house_number):
                        return number[1]
