import main
import creds
import requests


query_1 = (f"&api_key={creds.api_key}&address={main.customer.address}&tilt={main.array_1.tilt}&azimuth={main.array_1.azimuth}&losses={main.array_1.losses}&module_type={main.customer.module_type}array_type={main.customer.array_type}&system_capacity={main.system_capacity_1}")
query_2 = (f"&api_key={creds.api_key}&address={main.customer.address}&tilt={main.array_2.tilt}&azimuth={main.array_2.azimuth}&losses={main.array_2.losses}&module_type={main.customer.module_type}array_type={main.customer.array_type}&system_capacity={main.system_capacity_2}")
query_3 = (f"&api_key={creds.api_key}&address={main.customer.address}&tilt={main.array_3.tilt}&azimuth={main.array_3.azimuth}&losses={main.array_3.losses}&module_type={main.customer.module_type}array_type={main.customer.array_type}&system_capacity={main.system_capacity_3}")
query_4 = (f"&api_key={creds.api_key}&address={main.customer.address}&tilt={main.array_4.tilt}&azimuth={main.array_4.azimuth}&losses={main.array_4.losses}&module_type={main.customer.module_type}array_type={main.customer.array_type}&system_capacity={main.system_capacity_4}")
query_5 = (f"&api_key={creds.api_key}&address={main.customer.address}&tilt={main.array_5.tilt}&azimuth={main.array_5.azimuth}&losses={main.array_5.losses}&module_type={main.customer.module_type}array_type={main.customer.array_type}&system_capacity={main.system_capacity_5}")
query_6 = (f"&api_key={creds.api_key}&address={main.customer.address}&tilt={main.array_6.tilt}&azimuth={main.array_6.azimuth}&losses={main.array_6.losses}&module_type={main.customer.module_type}array_type={main.customer.array_type}&system_capacity={main.system_capacity_6}")

response_1 = requests.get("https://developer.nrel.gov/api/pvwatts/v6.json?" + query_1)
response_2 = requests.get("https://developer.nrel.gov/api/pvwatts/v6.json?" + query_2)
response_3 = requests.get("https://developer.nrel.gov/api/pvwatts/v6.json?" + query_3)
response_4 = requests.get("https://developer.nrel.gov/api/pvwatts/v6.json?" + query_4)
response_5 = requests.get("https://developer.nrel.gov/api/pvwatts/v6.json?" + query_5)
response_6 = requests.get("https://developer.nrel.gov/api/pvwatts/v6.json?" + query_6)