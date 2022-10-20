import creds
import gspread
import requests
import time
from docxtpl import DocxTemplate
import multiprocessing as mp

login = gspread.service_account(filename="service_account.json")
sheet_name = login.open("HOA")
worksheet = sheet_name.worksheet("10 Percent")
values = worksheet.get_values("B1:F35")
date = values[21][0]

ch = '.'

class Customer:
    def __init__(self, name, state, hoa_name, address, mod_watt, array_count, module_type, array_type):
        self.name = name
        self.state = state
        self.hoa_name = hoa_name
        self.address = address
        self.mod_watt = mod_watt
        self.array_count = array_count
        self.module_type = module_type
        self.array_type = array_type
        
customer = Customer(values[12][3], values[3][2], values[6][0], values[12][0].replace(" ", "%20"), float(values[9][4]), int(values[17][1]), 1, 1)

class Array_1:
    def __init__(self, tilt, azimuth, losses, quantity, direction):
        self.tilt = tilt
        self.azimuth = azimuth
        self.losses = losses
        self.quantity = quantity
        self.direction = direction
class Array_2:
    def __init__(self, tilt, azimuth, losses, quantity, direction):
        self.tilt = tilt
        self.azimuth = azimuth
        self.losses = losses
        self.quantity = quantity
        self.direction = direction
class Array_3:
    def __init__(self, tilt, azimuth, losses, quantity, direction):
        self.tilt = tilt
        self.azimuth = azimuth
        self.losses = losses
        self.quantity = quantity
        self.direction = direction
class Array_4:
    def __init__(self, tilt, azimuth, losses, quantity, direction):
        self.tilt = tilt
        self.azimuth = azimuth
        self.losses = losses
        self.quantity = quantity
        self.direction = direction
class Array_5:
    def __init__(self, tilt, azimuth, losses, quantity, direction):
        self.tilt = tilt
        self.azimuth = azimuth
        self.losses = losses
        self.quantity = quantity
        self.direction = direction
class Array_6:
    def __init__(self, tilt, azimuth, losses, quantity, direction):
        self.tilt = tilt
        self.azimuth = azimuth
        self.losses = losses
        self.quantity = quantity
        self.direction = direction

array_1 = Array_1(values[16][3], values[17][3], values[18][3], int(values[19][3]), values[20][3])
array_2 = Array_2(values[16][4], values[17][4], values[18][4], int(values[19][4]), values[20][4])
array_3 = Array_3(values[23][3], values[24][3], values[25][3], int(values[26][3]), values[27][3])
array_4 = Array_4(values[23][4], values[24][4], values[25][4], int(values[26][4]), values[27][4])
array_5 = Array_5(values[30][3], values[31][3], values[32][3], int(values[33][3]), values[34][3])
array_6 = Array_6(values[30][4], values[31][4], values[32][4], int(values[33][4]), values[34][4])

if customer.state == 'TX':
    customer.state = f"""
Here is a short excerpt from the Texas Solar Rights that refers to this issue. “The law also 
stipulates that the HOA may designate where the solar device should be located on a roof,
unless a homeowner can show that the designation negatively impacts the performance
of the solar energy device and an alternative location would increase production by
more than 10%. To show this, the law requires that modeling tools provided by the
National Renewable Laboratory (NREL) be used.” 

While not specified by name in the law, one of NREL’s available tools that can accomplish this is called PVWatts Calculator.
http://programs.dsireusa.org/system/program/detail/4880"""
elif customer.state == 'CO':
    customer.state = f"""
Here is a short excerpt from the Colorado House Bill that refers to this issue.
"Section 2 of the act adds specificity to the requirements that HOAs allow installation
of renewable energy generation devices (e.g solar panels) subject to reasonable
aesthetic guidelines by requiring approval or denial of a completed application
within 60 days and requiring approval if imposition of the aesthetic
guidelines would result in more than a 10% reduction in efficiency or a 10% increase in price

https://leg.colorado.gov/sites/default/files/2021a_1229_signed.pdf"""
else:
    pass

def letterOne():
    start_time = time.time()

    system_capacity_1 = customer.mod_watt * array_1.quantity
    system_capacity_2 = customer.mod_watt * array_2.quantity

    query_1 = (f"&api_key={creds.api_key}&address={customer.address}&tilt={array_1.tilt}&azimuth={array_1.azimuth}&losses={array_1.losses}&module_type={customer.module_type}&array_type={customer.array_type}&system_capacity={system_capacity_1}")
    query_2 = (f"&api_key={creds.api_key}&address={customer.address}&tilt={array_2.tilt}&azimuth={array_2.azimuth}&losses={array_2.losses}&module_type={customer.module_type}&array_type={customer.array_type}&system_capacity={system_capacity_2}")

    response_1 = requests.get("https://developer.nrel.gov/api/pvwatts/v6.json?" + query_1)
    response_2 = requests.get("https://developer.nrel.gov/api/pvwatts/v6.json?" + query_2)

    json_data_1 = (response_1.json())
    json_data_2 = (response_2.json())

    json_data_1 = int(json_data_1['outputs']['ac_annual'])
    json_data_2 = int(json_data_2['outputs']['ac_annual'])

    difference = int(json_data_1) - int(json_data_2)
    total = difference / int(json_data_1)

    doc = DocxTemplate("TEN_PERCENT_V5.docx")
    context = {'hoa_name': customer.hoa_name, 'date': date, 'name': customer.name,
               'quantity': array_1.quantity, 'old_direction': array_1.direction, 'quantity2': array_2.quantity, 'state': customer.state,
               'old_azimuth': array_1.azimuth.replace("azimuth=", ""), 'old_tilt': array_1.tilt.replace("tilt=", ""), 'new_direction': array_2.direction,
               'new_azimuth': array_2.azimuth.replace("azimuth=", ""), 'new_tilt': array_2.tilt.replace("tilt=", ""), 
               'mod_watt': str(customer.mod_watt).replace("0.", "").strip(), 'percent': str(total)[2:][:2] + "%",
               'ac_monthly_original': str(json_data_1).split(ch, 1)[0], 'ac_monthly_new': str(json_data_2).split(ch, 1)[0]}

    doc.render(context)
    doc.save(customer.name + " Ten Percent Letter 1.docx")

    doc = DocxTemplate("LETTER.docx")
    context = {'hoa_name': customer.hoa_name, 'date': date, 'name': customer.name,
               'quantity': array_1.quantity, 'old_direction': array_1.direction, 'quantity2': array_2.quantity, 'state': customer.state,
               'old_azimuth': array_1.azimuth.replace("azimuth=", ""), 'old_tilt': array_1.tilt.replace("tilt=", ""), 'new_direction': array_2.direction,
               'new_azimuth': array_2.azimuth.replace("azimuth=", ""), 'new_tilt': array_2.tilt.replace("tilt=", ""), 
               'mod_watt': str(customer.mod_watt).replace("0.", "").strip(), 'percent': str(total)[2:][:2] + "%",
               'ac_monthly_original': str(json_data_1).split(ch, 1)[0], 'ac_monthly_new': str(json_data_2).split(ch, 1)[0]}

    doc.render(context)
    doc.save(customer.name + " PVWatts Calculation.docx")

    end_time = time.time()
    final_time = end_time - start_time
    print(f"Letter One run in: {final_time} seconds")

def letterTwo():
    start_time = time.time()
    system_capacity_3 = customer.mod_watt * array_3.quantity
    system_capacity_4 = customer.mod_watt * array_4.quantity

    query_3 = (f"&api_key={creds.api_key}&address={customer.address}&tilt={array_3.tilt}&azimuth={array_3.azimuth}&losses={array_3.losses}&module_type={customer.module_type}&array_type={customer.array_type}&system_capacity={system_capacity_3}")
    query_4 = (f"&api_key={creds.api_key}&address={customer.address}&tilt={array_4.tilt}&azimuth={array_4.azimuth}&losses={array_4.losses}&module_type={customer.module_type}&array_type={customer.array_type}&system_capacity={system_capacity_4}")

    response_3 = requests.get("https://developer.nrel.gov/api/pvwatts/v6.json?" + query_3)
    response_4 = requests.get("https://developer.nrel.gov/api/pvwatts/v6.json?" + query_4)

    json_data_3 = (response_3.json())
    json_data_4 = (response_4.json())

    json_data_3 = int(json_data_3['outputs']['ac_annual'])
    json_data_4 = int(json_data_4['outputs']['ac_annual'])

    difference = int(json_data_3) - int(json_data_4)
    total = difference / int(json_data_3)

    doc = DocxTemplate("TEN_PERCENT_V5.docx")
    context = {'hoa_name': customer.hoa_name, 'date': date, 'name': customer.name,
               'quantity': array_3.quantity, 'old_direction': array_3.direction, 'quantity2': array_4.quantity, 'state': customer.state,
               'old_azimuth': array_3.azimuth.replace("azimuth=", ""), 'old_tilt': array_3.tilt.replace("tilt=", ""), 'new_direction': array_4.direction,
               'new_azimuth': array_4.azimuth.replace("azimuth=", ""), 'new_tilt': array_4.tilt.replace("tilt=", ""), 
               'mod_watt': str(customer.mod_watt).replace("0.", "").strip(), 'percent': str(total)[2:][:2] + "%",
               'ac_monthly_original': str(json_data_3).split(ch, 1)[0], 'ac_monthly_new': str(json_data_4).split(ch, 1)[0]}

    doc.render(context)
    doc.save(customer.name + " Ten Percent Letter 2.docx")

    end_time = time.time()
    final_time = end_time - start_time
    print(f"Letter Two run in: {final_time} seconds")

def letterThree():
    start_time = time.time()
    system_capacity_5 = customer.mod_watt * array_5.quantity
    system_capacity_6 = customer.mod_watt * array_6.quantity

    query_5 = (f"&api_key={creds.api_key}&address={customer.address}&tilt={array_5.tilt}&azimuth={array_5.azimuth}&losses={array_5.losses}&module_type={customer.module_type}&array_type={customer.array_type}&system_capacity={system_capacity_5}")
    query_6 = (f"&api_key={creds.api_key}&address={customer.address}&tilt={array_6.tilt}&azimuth={array_6.azimuth}&losses={array_6.losses}&module_type={customer.module_type}&array_type={customer.array_type}&system_capacity={system_capacity_6}")

    response_5 = requests.get("https://developer.nrel.gov/api/pvwatts/v6.json?" + query_5)
    response_6 = requests.get("https://developer.nrel.gov/api/pvwatts/v6.json?" + query_6)

    json_data_5 = (response_5.json())
    json_data_6 = (response_6.json())

    json_data_5 = int(json_data_5['outputs']['ac_annual'])
    json_data_6 = int(json_data_6['outputs']['ac_annual'])

    difference = int(json_data_5) - int(json_data_6)
    total = difference / int(json_data_5)

    doc = DocxTemplate("TEN_PERCENT_V5.docx")
    context = {'hoa_name': customer.hoa_name, 'date': date, 'name': customer.name,
               'quantity': array_5.quantity, 'old_direction': array_5.direction, 'quantity2': array_6.quantity, 'state': customer.state,
               'old_azimuth': array_5.azimuth.replace("azimuth=", ""), 'old_tilt': array_5.tilt.replace("tilt=", ""), 'new_direction': array_6.direction,
               'new_azimuth': array_6.azimuth.replace("azimuth=", ""), 'new_tilt': array_6.tilt.replace("tilt=", ""), 
               'mod_watt': str(customer.mod_watt).replace("0.", "").strip(), 'percent': str(total)[2:][:2] + "%",
               'ac_monthly_original': str(json_data_5).split(ch, 1)[0], 'ac_monthly_new': str(json_data_6).split(ch, 1)[0]}

    doc.render(context)
    doc.save(customer.name + " Ten Percent Letter 3.docx")

    end_time = time.time()
    final_time = end_time - start_time
    print(f"Letter Three run in: {final_time} seconds")

def main():
    if customer.array_count == 1:
        p1 = mp.Process(target=letterOne(), args=())
        p1.start()
    elif customer.array_count == 2:
        p1 = mp.Process(target=letterOne(), args=())
        p2 = mp.Process(target=letterTwo(), args=())
        p1.start()
        p2.start()
    elif customer.array_count == 3:
        p1 = mp.Process(target=letterOne(), args=())
        p2 = mp.Process(target=letterTwo(), args=())
        p3 = mp.Process(target=letterThree(), args=())
        p1.start()
        p2.start()
        p3.start()
    else:
        print(customer.array_count)
        exit()

if __name__ == '__main__':
    start_time = time.time()
    main()