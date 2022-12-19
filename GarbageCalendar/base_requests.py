import requests


class BaseRequests:
    base_url = "https://{}.jumomind.com/mmapp/api.php"
    endpoints = {
        "Ingolstadt": "ingol",
    }
    city_id = ""
    area_id = ""
    end_point = ""

    def __init__(self, city: str):
        if city not in self.endpoints.keys():
            raise ValueError(
                f"The city '{city}' is not yet supported. Open an issue and I will look into it."
            )
        self.end_point = self.endpoints[city]
        url = self.get_url_cities()
        response = requests.get(url)
        data = response.json()
        self.city_id = data[0]["id"]

    def get_url_cities(self):
        url = f"{self.base_url.format(self.end_point)}?r=cities"
        return url

    def get_url_streets(self):
        url = f"{self.base_url.format(self.end_point)}?r=streets&city_id={self.city_id}"
        return url

    def get_url_monthly_dates(self, year, month):
        url = "{}?r=calendar/{}-{:02d}&city_id={}&area_id={}".format(
            self.base_url.format(self.end_point),
            year,
            month,
            self.city_id,
            self.area_id,
        )
        return url
