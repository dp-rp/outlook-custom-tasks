# import http.client
# import json
# from pprint import pprint

# conn = http.client.HTTPSConnection("cost-of-living-and-prices.p.rapidapi.com")

# headers = {
#     'x-rapidapi-key': "65ddc0f62dmshfe17a963992b1a8p1cc800jsn07436358e68f",
#     'x-rapidapi-host': "cost-of-living-and-prices.p.rapidapi.com"
# }

# conn.request("GET", "/cities", headers=headers)

# res = conn.getresponse()
# data = res.read()

# cities_str = data.decode("utf-8")
# cities = json.loads(cities_str)

# pprint(cities.keys())
# pprint(cities['cities'].keys())

import json
from pprint import pprint

with open("./practical-portfolio-tracker/20241220-cost-of-living.json", 'r', encoding='utf-8') as f:    
    cost_of_living_cities = json.load(f)
    cities = cost_of_living_cities['cities']
    australian_cities = filter(lambda c: c['country_name'] == "Australia", cities)
    # pprint([c for c in australian_cities])
    melbourne = filter(lambda c: c['city_name'] == "Melbourne", australian_cities)
    pprint([c for c in melbourne])
