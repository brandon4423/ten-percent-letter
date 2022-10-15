import gspread
import creds
import requests
import time
from docxtpl import DocxTemplate

login = gspread.service_account(filename="service_account.json")
sheet_name = login.open("HOA")
tab_lookup = sheet_name.worksheet("10 Percent")
date = str(tab_lookup.acell("H7").value)

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

customer = Customer(str(tab_lookup.acell("D7").value), str(tab_lookup.acell("D4").value), str(tab_lookup.acell("B7").value), 
                    str(tab_lookup.acell("B13").value).replace(" ", "%20").strip(), float(tab_lookup.acell("F10").value), int(tab_lookup.acell("J2").value), "1", "1")

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

def arrayOne():
    start_time = time.time()

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

    array_1 = Array_1(str(tab_lookup.acell("M2").value), str(tab_lookup.acell("M3").value), str(tab_lookup.acell("M4").value), int(tab_lookup.acell("M5").value)
                    , str(tab_lookup.acell("M6").value))

    array_2 = Array_2(str(tab_lookup.acell("N2").value), str(tab_lookup.acell("N3").value), str(tab_lookup.acell("N4").value), int(tab_lookup.acell("N5").value)
                    , str(tab_lookup.acell("N6").value))

    system_capacity_1 = customer.mod_watt * array_1.quantity
    system_capacity_2 = customer.mod_watt * array_2.quantity

    api_param = "&api_key=" + creds.api_key
    old_query = (f"{api_param}&address={customer.address}&tilt={array_1.tilt}&azimuth={array_1.azimuth}&losses={array_1.losses}&module_type={customer.module_type}&array_type={customer.array_type}&system_capacity={system_capacity_1}")
    new_query = (f"{api_param}&address={customer.address}&tilt={array_2.tilt}&azimuth={array_2.azimuth}&losses={array_2.losses}&module_type={customer.module_type}&array_type={customer.array_type}&system_capacity={system_capacity_2}")

    response_1 = requests.get("https://developer.nrel.gov/api/pvwatts/v6.json?" + old_query)
    response_2 = requests.get("https://developer.nrel.gov/api/pvwatts/v6.json?" + new_query)

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

    end_time = time.time()
    final_time = end_time - start_time
    print(f"Letter one run in: {final_time} seconds")

def arrayTwo():
    start_time = time.time()

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

    array_1 = Array_1(str(tab_lookup.acell("M9").value), str(tab_lookup.acell("M10").value), str(tab_lookup.acell("M11").value), int(tab_lookup.acell("M12").value)
                    , str(tab_lookup.acell("M13").value))

    array_2 = Array_2(str(tab_lookup.acell("N9").value), str(tab_lookup.acell("N10").value), str(tab_lookup.acell("N11").value), int(tab_lookup.acell("N12").value)
                    , str(tab_lookup.acell("N13").value))

    system_capacity_1 = customer.mod_watt * array_1.quantity
    system_capacity_2 = customer.mod_watt * array_2.quantity

    api_param = "&api_key=" + creds.api_key
    old_query = (f"{api_param}&address={customer.address}&tilt={array_1.tilt}&azimuth={array_1.azimuth}&losses={array_1.losses}&module_type={customer.module_type}&array_type={customer.array_type}&system_capacity={system_capacity_1}")
    new_query = (f"{api_param}&address={customer.address}&tilt={array_2.tilt}&azimuth={array_2.azimuth}&losses={array_2.losses}&module_type={customer.module_type}&array_type={customer.array_type}&system_capacity={system_capacity_2}")

    response_1 = requests.get("https://developer.nrel.gov/api/pvwatts/v6.json?" + old_query)
    response_2 = requests.get("https://developer.nrel.gov/api/pvwatts/v6.json?" + new_query)

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
    doc.save(customer.name + " Ten Percent Letter 2.docx")

    end_time = time.time()
    final_time = end_time - start_time
    print(f"Letter Two run in: {final_time} seconds")

def arrayThree():
    start_time = time.time()

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

    array_1 = Array_1(str(tab_lookup.acell("M16").value), str(tab_lookup.acell("M17").value), str(tab_lookup.acell("M18").value), int(tab_lookup.acell("M19").value)
                    , str(tab_lookup.acell("M20").value))

    array_2 = Array_2(str(tab_lookup.acell("N16").value), str(tab_lookup.acell("N17").value), str(tab_lookup.acell("N18").value), int(tab_lookup.acell("N19").value)
                    , str(tab_lookup.acell("N20").value))

    system_capacity_1 = customer.mod_watt * array_1.quantity
    system_capacity_2 = customer.mod_watt * array_2.quantity

    api_param = "&api_key=" + creds.api_key
    old_query = (f"{api_param}&address={customer.address}&tilt={array_1.tilt}&azimuth={array_1.azimuth}&losses={array_1.losses}&module_type={customer.module_type}&array_type={customer.array_type}&system_capacity={system_capacity_1}")
    new_query = (f"{api_param}&address={customer.address}&tilt={array_2.tilt}&azimuth={array_2.azimuth}&losses={array_2.losses}&module_type={customer.module_type}&array_type={customer.array_type}&system_capacity={system_capacity_2}")

    response_1 = requests.get("https://developer.nrel.gov/api/pvwatts/v6.json?" + old_query)
    response_2 = requests.get("https://developer.nrel.gov/api/pvwatts/v6.json?" + new_query)

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
    doc.save(customer.name + " Ten Percent Letter 3.docx")

    end_time = time.time()
    final_time = end_time - start_time
    print(f"Letter Three run in: {final_time} seconds")

def main():
    if customer.array_count == 1:
        arrayOne()
    elif customer.array_count == 2:
        arrayOne(), arrayTwo()
    elif customer.array_count == 3:
        arrayOne(), arrayTwo(), arrayThree()
    else:
        exit()

if __name__ == '__main__':
    main()