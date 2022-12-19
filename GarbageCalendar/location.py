from typing import List, Optional

import requests

from .base_requests import BaseRequests


class Location(BaseRequests):
    def __init__(self, city: str):
        super().__init__(city)
        response = requests.get(self.get_url_streets())
        self.data = response.json()

    def get_street_house_numbers(self, street_name: str) -> List[str]:
        result = []
        for street in self.data:
            if street["name"] == street_name:
                for house_number in street["houseNumbers"]:
                    result.append(house_number[0])
                break
        return result

    def get_street_names_like(self, typed: Optional[str]) -> List[str]:
        result = []
        if typed is None or typed == "":
            return [street["name"] for street in self.data]
        for street in self.data:
            if street["name"].startswith(typed):
                result.append(street["name"])
        return result

    def get_ref_area(self, selected_street_name: str, selected_house_number: str) -> int:
        for street in self.data:
            if street["name"] == selected_street_name:
                for number in street["houseNumbers"]:
                    if number[0].__eq__(selected_house_number):
                        self.area_id = number[1]
                        return self.area_id
        return -1
