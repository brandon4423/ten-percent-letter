import gspread
import creds
import requests
import json
import time
from docxtpl import DocxTemplate

login = gspread.service_account(filename="service_account.json")
sheet_name = login.open("HOA")
tab_lookup = sheet_name.worksheet("10 Percent")
date = str(tab_lookup.acell("H7").value)

class customer:
    pass

customer = customer()
customer.state = str(tab_lookup.acell("D4").value)
customer.hoa_name = str(tab_lookup.acell("B7").value)
customer.name = str(tab_lookup.acell("D7").value)
customer.array_count = int(tab_lookup.acell("J2").value)
customer.address = str(tab_lookup.acell("B13").value)
customer.module_type = "1"
customer.array_type = "1"

customer.address = customer.address.replace(" ", "%20").strip()
customer.address = ("address=" + customer.address + "&")
customer.module_type = ("module_type=" + customer.module_type + "&")
customer.array_type = ("array_type=" + customer.array_type + "&")

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

    class array:
        pass

    array.mod_watt = float(tab_lookup.acell("F10").value)
    array.original_tilt = str(tab_lookup.acell("M2").value)
    array.new_tilt = str(tab_lookup.acell("N2").value)
    array.original_azimuth = str(tab_lookup.acell("M3").value)
    array.new_azimuth = str(tab_lookup.acell("N3").value)
    array.losses_original = str(tab_lookup.acell("M4").value)
    array.losses_new = str(tab_lookup.acell("N4").value)
    array.quantity = int(tab_lookup.acell("M5").value)
    array.quantity_2 = int(tab_lookup.acell("N5").value)
    array.original_direction = str(tab_lookup.acell("M6").value)
    array.new_direction = str(tab_lookup.acell("N6").value)
    array.system_capacity = array.mod_watt * array.quantity
    array.system_capacity_2 = array.mod_watt * array.quantity_2

    array.new_tilt = ("tilt=" + array.new_tilt + "&")
    array.new_azimuth = ("azimuth=" + array.new_azimuth + "&")
    array.original_tilt = ("tilt=" + array.original_tilt + "&")
    array.original_azimuth = ("azimuth=" + array.original_azimuth + "&")
    array.losses_original = ("losses=" + array.losses_original + "&")
    array.losses_new = ("losses=" + array.losses_new + "&")
    array.system_capacity = ("system_capacity=" + str(array.system_capacity) + "&")
    array.system_capacity_2 = ("system_capacity=" + str(array.system_capacity_2) + "&")

    api_param = "&api_key="
    old_query = customer.address + array.original_tilt + array.original_azimuth + array.losses_original + customer.module_type + customer.array_type + array.system_capacity + api_param + creds.api_key
    new_query = customer.address + array.new_tilt + array.new_azimuth + array.losses_new + customer.module_type + customer.array_type + array.system_capacity_2 + api_param + creds.api_key

    response = requests.get("https://developer.nrel.gov/api/pvwatts/v6.json?" + api_param + creds.api_key)
    base_url = "https://developer.nrel.gov/api/pvwatts/v6.json?"

    json_link_original = (base_url + old_query)
    json_link_new = (base_url + new_query)
    data_original = requests.get(json_link_original)
    data_new = requests.get(json_link_new)

    content = data_original.text
    data_original = json.loads(content)
    content = data_new.text
    data_new = json.loads(content)

    ac_monthly_original = data_original.get('outputs')
    ac_monthly_new = data_new.get('outputs')
    dict.items(ac_monthly_original)
    dict.items(ac_monthly_new)

    del [ac_monthly_original['ac_monthly']]
    del [ac_monthly_original['poa_monthly']]
    del [ac_monthly_original['solrad_monthly']]
    del [ac_monthly_original['dc_monthly']]
    del [ac_monthly_original['solrad_annual']]
    del [ac_monthly_original['capacity_factor']]
    ac_monthly_original = str(ac_monthly_original)
    ac_monthly_original = ac_monthly_original[14:]

    del [ac_monthly_new['ac_monthly']]
    del [ac_monthly_new['poa_monthly']]
    del [ac_monthly_new['solrad_monthly']]
    del [ac_monthly_new['dc_monthly']]
    del [ac_monthly_new['solrad_annual']]
    del [ac_monthly_new['capacity_factor']]
    ac_monthly_new = str(ac_monthly_new)
    ac_monthly_new = ac_monthly_new[14:]

    ch = '.'
    try:
        ac_monthly_original = ac_monthly_original[0: ac_monthly_original.index(ch)]
    except ValueError:
        pass

    ch = '.'
    try:
        ac_monthly_new = ac_monthly_new[0: ac_monthly_new.index(ch)]
    except ValueError:
        pass

    ac_monthly_original = int(ac_monthly_original)
    ac_monthly_new = int(ac_monthly_new)

    difference = ac_monthly_original - ac_monthly_new
    total = difference / ac_monthly_original

    if total <= 0.1:
        total = str(total)
        total = total[2:]
        total = total[:2]
        total = total[1:]
    else:
        total = str(total)
        total = total[2:]
        total = total[:2]

    total = str(total)
    total = total + "%"

    array.original_tilt = str(tab_lookup.acell("M2").value)
    array.new_tilt = str(tab_lookup.acell("N2").value)
    array.original_azimuth = str(tab_lookup.acell("M3").value)
    array.new_azimuth = str(tab_lookup.acell("N3").value)
    array.mod_watt = str(array.mod_watt)
    array.mod_watt = array.mod_watt.replace("0.", "").strip()

    doc = DocxTemplate("TEN_PERCENT_V5.docx")
    context = {'hoa_name': customer.hoa_name, 'date': date, 'name': customer.name,
               'quantity': array.quantity, 'old_direction': array.original_direction, 'quantity2': array.quantity_2, 'state': customer.state,
               'old_azimuth': array.original_azimuth, 'old_tilt': array.original_tilt, 'new_direction': array.new_direction,
               'new_azimuth': array.new_azimuth, 'new_tilt': array.new_tilt, 'mod_watt': array.mod_watt, 'percent': total,
               'ac_monthly_original': ac_monthly_original, 'ac_monthly_new': ac_monthly_new}

    doc.render(context)
    doc.save(customer.name + " Ten Percent Letter 1.docx")

    end_time = time.time()
    final_time = end_time - start_time
    print(f"Letter one run in: {final_time} seconds")

def arrayTwo():
    start_time = time.time()

    class array:
        pass

    array.mod_watt = float(tab_lookup.acell("F10").value)
    array.original_tilt = str(tab_lookup.acell("M9").value)
    array.new_tilt = str(tab_lookup.acell("N9").value)
    array.original_azimuth = str(tab_lookup.acell("M10").value)
    array.new_azimuth = str(tab_lookup.acell("N10").value)
    array.losses_original = str(tab_lookup.acell("M11").value)
    array.losses_new = str(tab_lookup.acell("N11").value)
    array.quantity = int(tab_lookup.acell("M12").value)
    array.quantity_2 = int(tab_lookup.acell("N12").value)
    array.original_direction = str(tab_lookup.acell("M13").value)
    array.new_direction = str(tab_lookup.acell("N13").value)
    array.system_capacity = array.mod_watt * array.quantity
    array.system_capacity_2 = array.mod_watt * array.quantity_2

    array.new_tilt = ("tilt=" + array.new_tilt + "&")
    array.new_azimuth = ("azimuth=" + array.new_azimuth + "&")
    array.original_tilt = ("tilt=" + array.original_tilt + "&")
    array.original_azimuth = ("azimuth=" + array.original_azimuth + "&")
    array.losses_original = ("losses=" + array.losses_original + "&")
    array.losses_new = ("losses=" + array.losses_new + "&")
    array.system_capacity = ("system_capacity=" + str(array.system_capacity) + "&")
    array.system_capacity_2 = ("system_capacity=" + str(array.system_capacity_2) + "&")

    api_param = "&api_key="
    old_query = customer.address + array.original_tilt + array.original_azimuth + array.losses_original + customer.module_type + customer.array_type + array.system_capacity + api_param + creds.api_key
    new_query = customer.address + array.new_tilt + array.new_azimuth + array.losses_new + customer.module_type + customer.array_type + array.system_capacity_2 + api_param + creds.api_key

    response = requests.get("https://developer.nrel.gov/api/pvwatts/v6.json?" + api_param + creds.api_key)
    base_url = "https://developer.nrel.gov/api/pvwatts/v6.json?"

    json_link_original = (base_url + old_query)
    json_link_new = (base_url + new_query)
    data_original = requests.get(json_link_original)
    data_new = requests.get(json_link_new)

    content = data_original.text
    data_original = json.loads(content)
    content = data_new.text
    data_new = json.loads(content)

    ac_monthly_original = data_original.get('outputs')
    ac_monthly_new = data_new.get('outputs')
    dict.items(ac_monthly_original)
    dict.items(ac_monthly_new)

    del [ac_monthly_original['ac_monthly']]
    del [ac_monthly_original['poa_monthly']]
    del [ac_monthly_original['solrad_monthly']]
    del [ac_monthly_original['dc_monthly']]
    del [ac_monthly_original['solrad_annual']]
    del [ac_monthly_original['capacity_factor']]
    ac_monthly_original = str(ac_monthly_original)
    ac_monthly_original = ac_monthly_original[14:]

    del [ac_monthly_new['ac_monthly']]
    del [ac_monthly_new['poa_monthly']]
    del [ac_monthly_new['solrad_monthly']]
    del [ac_monthly_new['dc_monthly']]
    del [ac_monthly_new['solrad_annual']]
    del [ac_monthly_new['capacity_factor']]
    ac_monthly_new = str(ac_monthly_new)
    ac_monthly_new = ac_monthly_new[14:]

    ch = '.'
    try:
        ac_monthly_original = ac_monthly_original[0: ac_monthly_original.index(ch)]
    except ValueError:
        pass

    ch = '.'
    try:
        ac_monthly_new = ac_monthly_new[0: ac_monthly_new.index(ch)]
    except ValueError:
        pass

    ac_monthly_original = int(ac_monthly_original)
    ac_monthly_new = int(ac_monthly_new)

    difference = ac_monthly_original - ac_monthly_new
    total = difference / ac_monthly_original

    if total <= 0.1:
        total = str(total)
        total = total[2:]
        total = total[:2]
        total = total[1:]
    else:
        total = str(total)
        total = total[2:]
        total = total[:2]

    total = str(total)
    total = total + "%"

    array.original_tilt = str(tab_lookup.acell("M9").value)
    array.new_tilt = str(tab_lookup.acell("N9").value)
    array.original_azimuth = str(tab_lookup.acell("M10").value)
    array.new_azimuth = str(tab_lookup.acell("N10").value)
    array.mod_watt = str(array.mod_watt)
    array.mod_watt = array.mod_watt.replace("0.", "").strip()

    doc = DocxTemplate("TEN_PERCENT_V5.docx")
    context = {'hoa_name': customer.hoa_name, 'date': date, 'name': customer.name,
               'quantity': array.quantity, 'old_direction': array.original_direction, 'quantity2': array.quantity_2, 'state': customer.state,
               'old_azimuth': array.original_azimuth, 'old_tilt': array.original_tilt, 'new_direction': array.new_direction,
               'new_azimuth': array.new_azimuth, 'new_tilt': array.new_tilt, 'mod_watt': array.mod_watt, 'percent': total,
               'ac_monthly_original': ac_monthly_original, 'ac_monthly_new': ac_monthly_new}

    doc.render(context)
    doc.save(customer.name + " Ten Percent Letter 2.docx")

    end_time = time.time()
    final_time = end_time - start_time
    print(f"Letter Two run in: {final_time} seconds")

def arrayThree():
    start_time = time.time()

    class array:
        pass

    array.mod_watt = float(tab_lookup.acell("F10").value)
    array.original_tilt = str(tab_lookup.acell("M16").value)
    array.new_tilt = str(tab_lookup.acell("N16").value)
    array.original_azimuth = str(tab_lookup.acell("M17").value)
    array.new_azimuth = str(tab_lookup.acell("N17").value)
    array.losses_original = str(tab_lookup.acell("M18").value)
    array.losses_new = str(tab_lookup.acell("N18").value)
    array.quantity = int(tab_lookup.acell("M19").value)
    array.quantity_2 = int(tab_lookup.acell("N19").value)
    array.original_direction = str(tab_lookup.acell("M20").value)
    array.new_direction = str(tab_lookup.acell("N20").value)
    array.system_capacity = array.mod_watt * array.quantity
    array.system_capacity_2 = array.mod_watt * array.quantity_2

    array.new_tilt = ("tilt=" + array.new_tilt + "&")
    array.new_azimuth = ("azimuth=" + array.new_azimuth + "&")
    array.original_tilt = ("tilt=" + array.original_tilt + "&")
    array.original_azimuth = ("azimuth=" + array.original_azimuth + "&")
    array.losses_original = ("losses=" + array.losses_original + "&")
    array.losses_new = ("losses=" + array.losses_new + "&")
    array.system_capacity = ("system_capacity=" + str(array.system_capacity) + "&")
    array.system_capacity_2 = ("system_capacity=" + str(array.system_capacity_2) + "&")

    api_param = "&api_key="
    old_query = customer.address + array.original_tilt + array.original_azimuth + array.losses_original + customer.module_type + customer.array_type + array.system_capacity + api_param + creds.api_key
    new_query = customer.address + array.new_tilt + array.new_azimuth + array.losses_new + customer.module_type + customer.array_type + array.system_capacity_2 + api_param + creds.api_key

    response = requests.get("https://developer.nrel.gov/api/pvwatts/v6.json?" + api_param + creds.api_key)
    base_url = "https://developer.nrel.gov/api/pvwatts/v6.json?"

    json_link_original = (base_url + old_query)
    json_link_new = (base_url + new_query)
    data_original = requests.get(json_link_original)
    data_new = requests.get(json_link_new)

    content = data_original.text
    data_original = json.loads(content)
    content = data_new.text
    data_new = json.loads(content)

    ac_monthly_original = data_original.get('outputs')
    ac_monthly_new = data_new.get('outputs')
    dict.items(ac_monthly_original)
    dict.items(ac_monthly_new)

    del [ac_monthly_original['ac_monthly']]
    del [ac_monthly_original['poa_monthly']]
    del [ac_monthly_original['solrad_monthly']]
    del [ac_monthly_original['dc_monthly']]
    del [ac_monthly_original['solrad_annual']]
    del [ac_monthly_original['capacity_factor']]
    ac_monthly_original = str(ac_monthly_original)
    ac_monthly_original = ac_monthly_original[14:]

    del [ac_monthly_new['ac_monthly']]
    del [ac_monthly_new['poa_monthly']]
    del [ac_monthly_new['solrad_monthly']]
    del [ac_monthly_new['dc_monthly']]
    del [ac_monthly_new['solrad_annual']]
    del [ac_monthly_new['capacity_factor']]
    ac_monthly_new = str(ac_monthly_new)
    ac_monthly_new = ac_monthly_new[14:]

    ch = '.'
    try:
        ac_monthly_original = ac_monthly_original[0: ac_monthly_original.index(ch)]
    except ValueError:
        pass

    ch = '.'
    try:
        ac_monthly_new = ac_monthly_new[0: ac_monthly_new.index(ch)]
    except ValueError:
        pass

    ac_monthly_original = int(ac_monthly_original)
    ac_monthly_new = int(ac_monthly_new)

    difference = ac_monthly_original - ac_monthly_new
    total = difference / ac_monthly_original

    if total <= 0.1:
        total = str(total)
        total = total[2:]
        total = total[:2]
        total = total[1:]
    else:
        total = str(total)
        total = total[2:]
        total = total[:2]

    total = str(total)
    total = total + "%"

    array.original_tilt = str(tab_lookup.acell("M16").value)
    array.new_tilt = str(tab_lookup.acell("N16").value)
    array.original_azimuth = str(tab_lookup.acell("M17").value)
    array.new_azimuth = str(tab_lookup.acell("N17").value)
    array.mod_watt = str(array.mod_watt)
    array.mod_watt = array.mod_watt.replace("0.", "").strip()

    doc = DocxTemplate("TEN_PERCENT_V5.docx")
    context = {'hoa_name': customer.hoa_name, 'date': date, 'name': customer.name,
               'quantity': array.quantity, 'old_direction': array.original_direction, 'quantity2': array.quantity_2, 'state': customer.state,
               'old_azimuth': array.original_azimuth, 'old_tilt': array.original_tilt, 'new_direction': array.new_direction,
               'new_azimuth': array.new_azimuth, 'new_tilt': array.new_tilt, 'mod_watt': array.mod_watt, 'percent': total,
               'ac_monthly_original': ac_monthly_original, 'ac_monthly_new': ac_monthly_new}

    doc.render(context)
    doc.save(customer.name + " Ten Percent Letter 3.docx")

    end_time = time.time()
    final_time = end_time - start_time
    print(f"Letter Three run in: {final_time} seconds")

def main():
    if customer.array_count == 1:
        arrayOne()
    elif customer.array_count == 2:
        arrayOne()
        arrayTwo()
    elif customer.array_count == 3:
        arrayOne()
        arrayTwo()
        arrayThree()
    else:
        exit()

if __name__ == '__main__':
    main()