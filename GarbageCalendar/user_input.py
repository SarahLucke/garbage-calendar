from .location import Location


class UserInput:
    def __init__(self, city: str, street: str, house_number: str, year: int):
        self.location = Location(city)
        self.location.get_ref_area(street, house_number)
        self.city = city
        self.street = street
        self.house_number = house_number
        self.year = year
