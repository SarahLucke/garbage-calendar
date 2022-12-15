import requests


class BaseRequests:
    base_url = "https://{}.jumomind.com/mmapp/api.php"
    endpoints = {
        "Ingolstadt": "ingol",
    }
    city_id = ""
    area_id = ""

    def __init__(self, city: str):
        if city not in self.endpoints.keys():
            raise ValueError(
                f"The city '{city}' is not yet supported. Open an issue and I will look into it."
            )
        self.endpoint = self.endpoints[city]
        url = self.get_url_cities()
        response = requests.get(url)
        data = response.json()
        self.city_id = data[0]["id"]

    def get_url_cities(self):
        url = self.base_url.format(self.endpoint) + "?r=cities"
        return url

    def get_url_streets(self):
        url = self.base_url.format(self.endpoint) + "?r=streets&city_id={}".format(
            self.city_id
        )
        return url

    def get_url_monthly_dates(self, year, month):
        url = self.base_url.format(
            self.endpoint
        ) + "?r=calendar/{}-{:02d}&city_id={}&area_id={}".format(
            year, month, self.city_id, self.area_id
        )
        return url
