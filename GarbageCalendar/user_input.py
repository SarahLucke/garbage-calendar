from .location import Location


class UserInput:
    city_id = ""
    area_id = ""

    def __init__(self, city: str, street: str, house_number: str):
        location = Location(city)
        self.city_id = location.city_id
        self.area_id = location.get_ref_area(street, house_number)
