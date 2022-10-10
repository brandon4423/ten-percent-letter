import gspread
import creds
import requests
import json
from docxtpl import DocxTemplate

#Login to google account services - pass api key from json file to connect python to google sheet ********
login = gspread.service_account(filename="service_account.json")
sheet_name = login.open("HOA")

tab_lookup = sheet_name.worksheet("10 Percent")
state = str(tab_lookup.acell("D4").value)
hoa_name = str(tab_lookup.acell("B7").value)
date = str(tab_lookup.acell("H7").value)
name = str(tab_lookup.acell("D7").value)
arrayCount = str(tab_lookup.acell("J2").value)
arrayCount = int(arrayCount)

if state == 'TX':
    state = f"""
Here is a short excerpt from the Texas Solar Rights that refers to this issue. “The law also 
stipulates that the HOA may designate where the solar device should be located on a roof,
unless a homeowner can show that the designation negatively impacts the performance
of the solar energy device and an alternative location would increase production by
more than 10%. To show this, the law requires that modeling tools provided by the
National Renewable Laboratory (NREL) be used.” 

While not specified by name in the law, one of NREL’s available tools that can accomplish this is called PVWatts Calculator.
http://programs.dsireusa.org/system/program/detail/4880"""
elif state == 'CO':
    state = f"""
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

    quantity = str(tab_lookup.acell("M5").value)
    quantity_2 = str(tab_lookup.acell("N5").value)
    old_tilt = str(tab_lookup.acell("M2").value)
    old_azimuth = str(tab_lookup.acell("M3").value)
    old_direction = str(tab_lookup.acell("M6").value)
    new_direction = str(tab_lookup.acell("N6").value)
    new_tilt = str(tab_lookup.acell("N2").value)
    new_azimuth = str(tab_lookup.acell("N3").value)
    mod_watt = str(tab_lookup.acell("C10").value)
    address = str(tab_lookup.acell("B13:C13").value)
    losses_o = str(tab_lookup.acell("M4").value)
    losses_n = str(tab_lookup.acell("N4").value)
    module_type = "1"
    array_type = "1"
    system_capacity = 1
    system_capacity_2 = 1
    # ^^^ Google sheet values - checked and stringed ready to pass into docxtpl and calculations ^^^ *********

    # Calculating system_capacity ****************************************************************************
    quantity = int(quantity)
    quantity_2 = int(quantity_2)

    if mod_watt == "SPR-M435-H-AC":
        system_capacity = quantity * .435
        system_capacity_2 = quantity_2 * .435
    elif mod_watt == "SPR-M425-H-AC":
        system_capacity = quantity * .425
        system_capacity_2 = quantity_2 * .425
    elif mod_watt == "SPR-A420-AC":
        system_capacity = quantity * .420
        system_capacity_2 = quantity_2 * .420
    elif mod_watt == "SPR-A415-AC":
        system_capacity = quantity * .415
        system_capacity_2 = quantity_2 * .415
    elif mod_watt == "SPR-A410-AC":
        system_capacity = quantity * .410
        system_capacity_2 = quantity_2 * .410
    elif mod_watt == "JKM410M-72HL-V G2 410W":
        system_capacity = quantity * .410
        system_capacity_2 = quantity_2 * .410
    elif mod_watt == "SPR-A400-BLK-AC":
        system_capacity = quantity * .400
        system_capacity_2 = quantity_2 * .400
    elif mod_watt == "SPR-A400-BLK":
        system_capacity = quantity * .400
        system_capacity_2 = quantity_2 * .400
    elif mod_watt == "SPR-A400-AC":
        system_capacity = quantity * .400
        system_capacity_2 = quantity_2 * .400
    elif mod_watt == "SPR-U400-BLK":
        system_capacity = quantity * .400
        system_capacity_2 = quantity_2 * .400
    elif mod_watt == "SPR-X22-370-AC":
        system_capacity = quantity * .370
        system_capacity_2 = quantity_2 * .370
    elif mod_watt == "SPR-X22-360-AC":
        system_capacity = quantity * .360
        system_capacity_2 = quantity_2 * .360
    elif mod_watt == "SPR-X22-360":
        system_capacity = quantity * .360
        system_capacity_2 = quantity_2 * .360
    elif mod_watt == "SPR-X21-350-BLK-AC":
        system_capacity = quantity * .350
        system_capacity_2 = quantity_2 * .350
    elif mod_watt == "SPR-E20-327-AC":
        system_capacity = quantity * .327
        system_capacity_2 = quantity_2 * .327
    elif mod_watt == "SPR-E20-327":
        system_capacity = quantity * .327
        system_capacity_2 = quantity_2 * .327
    elif mod_watt == "SPR-E19-320-AC":
        system_capacity = quantity * .320
        system_capacity_2 = quantity_2 * .320
    else:
        pass

    system_capacity = str(system_capacity)
    system_capacity_2 = str(system_capacity_2)

    # *********************************************************************************************************

    # setting string variables for NREL/PVWATTS parameters ****************************************************

    address = address.replace(" ", "%20").strip()
    address = ("address=" + address + "&")
    new_tilt = ("tilt=" + new_tilt + "&")
    new_azimuth = ("azimuth=" + new_azimuth + "&")
    old_tilt = ("tilt=" + old_tilt + "&")
    old_azimuth = ("azimuth=" + old_azimuth + "&")
    losses_o = ("losses=" + losses_o + "&")
    losses_n = ("losses=" + losses_n + "&")
    module_type = ("module_type=" + module_type + "&")
    array_type = ("array_type=" + array_type + "&")
    system_capacity = ("system_capacity=" + system_capacity + "&")
    system_capacity_2 = ("system_capacity=" + system_capacity_2 + "&")

    api_param = "&api_key="
    old_query = address + old_tilt + old_azimuth + losses_o + module_type + array_type + system_capacity + api_param + creds.api_key
    new_query = address + new_tilt + new_azimuth + losses_n + module_type + array_type + system_capacity_2 + api_param + creds.api_key

    # parameters set for NREL/PVW API connection, preparing to make get call & parse data after 200 response *

    # preforming first requests to NREL/PVW API -- Looks a little jank but it works **************************
    response = requests.get("https://developer.nrel.gov/api/pvwatts/v6.json?" + api_param + creds.api_key)
    base_url = "https://developer.nrel.gov/api/pvwatts/v6.json?"

    json_link_original = (base_url + old_query)
    json_link_new = (base_url + new_query)
    data_original = requests.get(json_link_original)
    data_new = requests.get(json_link_new)

    print(json_link_original)
    print(json_link_new)
    print(data_original.status_code)
    print(data_new.status_code)

    # finished api request should have 200 response in console / printed *************************************

    # had a lot of trouble with the json formatting from NREL/PVW had to parse this way to get annual total **
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
        # Remove all characters after the character '.' from string
        ac_monthly_original = ac_monthly_original[0: ac_monthly_original.index(ch)]
    except ValueError:
        pass

    ch = '.'
    try:
        # Remove all characters after the character '.' from string
        ac_monthly_new = ac_monthly_new[0: ac_monthly_new.index(ch)]
    except ValueError:
        pass

    # ^^^^ finished parsing / quite the mess but works great, on to calculate percent difference between calls

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

    # calculations for ten percent docx sheet finished, some parsing of useless string data ******************

    # setting variables and values for ten percent docx and finishing out with a final print *****************
    # This is here to fix the TEN_PERCENT letter back to original formatting
    old_tilt = str(tab_lookup.acell("M2").value)
    old_azimuth = str(tab_lookup.acell("M3").value)
    new_tilt = str(tab_lookup.acell("N2").value)
    new_azimuth = str(tab_lookup.acell("N3").value)
    line_break = "________________________________________________"

    doc = DocxTemplate("TEN_PERCENT_V5.docx")
    context = {'hoa_name': hoa_name, 'date': date, 'name': name,
               'quantity': quantity, 'old_direction': old_direction, 'quantity2': quantity_2, 'state': state,
               'old_azimuth': old_azimuth, 'old_tilt': old_tilt, 'new_direction': new_direction,
               'new_azimuth': new_azimuth, 'new_tilt': new_tilt, 'mod_watt': mod_watt, 'percent': total,
               'ac_monthly_original': ac_monthly_original, 'ac_monthly_new': ac_monthly_new}

    doc.render(context)
    doc.save(name + " Ten Percent Letter array 1.docx")
    print("Ten Percent Letter finished...")

def arrayTwo():

    quantity = str(tab_lookup.acell("M12").value)
    quantity_2 = str(tab_lookup.acell("N12").value)
    old_tilt = str(tab_lookup.acell("M9").value)
    old_azimuth = str(tab_lookup.acell("M10").value)
    old_direction = str(tab_lookup.acell("M13").value)
    new_direction = str(tab_lookup.acell("N13").value)
    new_tilt = str(tab_lookup.acell("N9").value)
    new_azimuth = str(tab_lookup.acell("N10").value)
    mod_watt = str(tab_lookup.acell("C10").value)
    address = str(tab_lookup.acell("B13:C13").value)
    losses_o = str(tab_lookup.acell("M11").value)
    losses_n = str(tab_lookup.acell("N11").value)
    module_type = "1"
    array_type = "1"
    system_capacity = 1
    system_capacity_2 = 1
    # ^^^ Google sheet values - checked and stringed ready to pass into docxtpl and calculations ^^^ *********

    # Calculating system_capacity ****************************************************************************
    quantity = int(quantity)
    quantity_2 = int(quantity_2)

    if mod_watt == "SPR-M435-H-AC":
        system_capacity = quantity * .435
        system_capacity_2 = quantity_2 * .435
    elif mod_watt == "SPR-M425-H-AC":
        system_capacity = quantity * .425
        system_capacity_2 = quantity_2 * .425
    elif mod_watt == "SPR-A420-AC":
        system_capacity = quantity * .420
        system_capacity_2 = quantity_2 * .420
    elif mod_watt == "SPR-A415-AC":
        system_capacity = quantity * .415
        system_capacity_2 = quantity_2 * .415
    elif mod_watt == "SPR-A410-AC":
        system_capacity = quantity * .410
        system_capacity_2 = quantity_2 * .410
    elif mod_watt == "JKM410M-72HL-V G2 410W":
        system_capacity = quantity * .410
        system_capacity_2 = quantity_2 * .410
    elif mod_watt == "SPR-A400-BLK-AC":
        system_capacity = quantity * .400
        system_capacity_2 = quantity_2 * .400
    elif mod_watt == "SPR-A400-BLK":
        system_capacity = quantity * .400
        system_capacity_2 = quantity_2 * .400
    elif mod_watt == "SPR-A400-AC":
        system_capacity = quantity * .400
        system_capacity_2 = quantity_2 * .400
    elif mod_watt == "SPR-U400-BLK":
        system_capacity = quantity * .400
        system_capacity_2 = quantity_2 * .400
    elif mod_watt == "SPR-X22-370-AC":
        system_capacity = quantity * .370
        system_capacity_2 = quantity_2 * .370
    elif mod_watt == "SPR-X22-360-AC":
        system_capacity = quantity * .360
        system_capacity_2 = quantity_2 * .360
    elif mod_watt == "SPR-X22-360":
        system_capacity = quantity * .360
        system_capacity_2 = quantity_2 * .360
    elif mod_watt == "SPR-X21-350-BLK-AC":
        system_capacity = quantity * .350
        system_capacity_2 = quantity_2 * .350
    elif mod_watt == "SPR-E20-327-AC":
        system_capacity = quantity * .327
        system_capacity_2 = quantity_2 * .327
    elif mod_watt == "SPR-E20-327":
        system_capacity = quantity * .327
        system_capacity_2 = quantity_2 * .327
    elif mod_watt == "SPR-E19-320-AC":
        system_capacity = quantity * .320
        system_capacity_2 = quantity_2 * .320
    else:
        pass

    system_capacity = str(system_capacity)
    system_capacity_2 = str(system_capacity_2)

    # *********************************************************************************************************

    # setting string variables for NREL/PVWATTS parameters ****************************************************

    address = address.replace(" ", "%20").strip()
    address = ("address=" + address + "&")
    new_tilt = ("tilt=" + new_tilt + "&")
    new_azimuth = ("azimuth=" + new_azimuth + "&")
    old_tilt = ("tilt=" + old_tilt + "&")
    old_azimuth = ("azimuth=" + old_azimuth + "&")
    losses_o = ("losses=" + losses_o + "&")
    losses_n = ("losses=" + losses_n + "&")
    module_type = ("module_type=" + module_type + "&")
    array_type = ("array_type=" + array_type + "&")
    system_capacity = ("system_capacity=" + system_capacity + "&")
    system_capacity_2 = ("system_capacity=" + system_capacity_2 + "&")

    api_param = "&api_key="
    old_query = address + old_tilt + old_azimuth + losses_o + module_type + array_type + system_capacity + api_param + creds.api_key
    new_query = address + new_tilt + new_azimuth + losses_n + module_type + array_type + system_capacity_2 + api_param + creds.api_key

    # parameters set for NREL/PVW API connection, preparing to make get call & parse data after 200 response *

    # preforming first requests to NREL/PVW API -- Looks a little jank but it works **************************
    response = requests.get("https://developer.nrel.gov/api/pvwatts/v6.json?" + api_param + creds.api_key)
    base_url = "https://developer.nrel.gov/api/pvwatts/v6.json?"

    json_link_original = (base_url + old_query)
    json_link_new = (base_url + new_query)
    data_original = requests.get(json_link_original)
    data_new = requests.get(json_link_new)

    print(json_link_original)
    print(json_link_new)
    print(data_original.status_code)
    print(data_new.status_code)

    # finished api request should have 200 response in console / printed *************************************

    # had a lot of trouble with the json formatting from NREL/PVW had to parse this way to get annual total **
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
        # Remove all characters after the character '.' from string
        ac_monthly_original = ac_monthly_original[0: ac_monthly_original.index(ch)]
    except ValueError:
        pass

    ch = '.'
    try:
        # Remove all characters after the character '.' from string
        ac_monthly_new = ac_monthly_new[0: ac_monthly_new.index(ch)]
    except ValueError:
        pass

    # ^^^^ finished parsing / quite the mess but works great, on to calculate percent difference between calls

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

    # calculations for ten percent docx sheet finished, some parsing of useless string data ******************

    # setting variables and values for ten percent docx and finishing out with a final print *****************
    # This is here to fix the TEN_PERCENT letter back to original formatting
    old_tilt = str(tab_lookup.acell("M9").value)
    old_azimuth = str(tab_lookup.acell("M10").value)
    new_tilt = str(tab_lookup.acell("N9").value)
    new_azimuth = str(tab_lookup.acell("N10").value)
    line_break = "________________________________________________"

    doc = DocxTemplate("TEN_PERCENT_V5.docx")
    context = {'hoa_name': hoa_name, 'date': date, 'name': name,
               'quantity': quantity, 'old_direction': old_direction, 'quantity2': quantity_2, 'state': state,
               'old_azimuth': old_azimuth, 'old_tilt': old_tilt, 'new_direction': new_direction,
               'new_azimuth': new_azimuth, 'new_tilt': new_tilt, 'mod_watt': mod_watt, 'percent': total,
               'ac_monthly_original': ac_monthly_original, 'ac_monthly_new': ac_monthly_new}

    doc.render(context)
    doc.save(name + " Ten Percent Letter array 2.docx")
    print("Ten Percent Letter finished...")

def arrayThree():

    quantity = str(tab_lookup.acell("M19").value)
    quantity_2 = str(tab_lookup.acell("N19").value)
    old_tilt = str(tab_lookup.acell("M16").value)
    old_azimuth = str(tab_lookup.acell("M17").value)
    old_direction = str(tab_lookup.acell("M20").value)
    new_direction = str(tab_lookup.acell("N20").value)
    new_tilt = str(tab_lookup.acell("N16").value)
    new_azimuth = str(tab_lookup.acell("N17").value)
    mod_watt = str(tab_lookup.acell("C10").value)
    address = str(tab_lookup.acell("B13:C13").value)
    losses_o = str(tab_lookup.acell("M18").value)
    losses_n = str(tab_lookup.acell("N18").value)
    module_type = "1"
    array_type = "1"
    system_capacity = 1
    system_capacity_2 = 1
    # ^^^ Google sheet values - checked and stringed ready to pass into docxtpl and calculations ^^^ *********

    # Calculating system_capacity ****************************************************************************
    quantity = int(quantity)
    quantity_2 = int(quantity_2)

    if mod_watt == "SPR-M435-H-AC":
        system_capacity = quantity * .435
        system_capacity_2 = quantity_2 * .435
    elif mod_watt == "SPR-M425-H-AC":
        system_capacity = quantity * .425
        system_capacity_2 = quantity_2 * .425
    elif mod_watt == "SPR-A420-AC":
        system_capacity = quantity * .420
        system_capacity_2 = quantity_2 * .420
    elif mod_watt == "SPR-A415-AC":
        system_capacity = quantity * .415
        system_capacity_2 = quantity_2 * .415
    elif mod_watt == "SPR-A410-AC":
        system_capacity = quantity * .410
        system_capacity_2 = quantity_2 * .410
    elif mod_watt == "JKM410M-72HL-V G2 410W":
        system_capacity = quantity * .410
        system_capacity_2 = quantity_2 * .410
    elif mod_watt == "SPR-A400-BLK-AC":
        system_capacity = quantity * .400
        system_capacity_2 = quantity_2 * .400
    elif mod_watt == "SPR-A400-BLK":
        system_capacity = quantity * .400
        system_capacity_2 = quantity_2 * .400
    elif mod_watt == "SPR-A400-AC":
        system_capacity = quantity * .400
        system_capacity_2 = quantity_2 * .400
    elif mod_watt == "SPR-U400-BLK":
        system_capacity = quantity * .400
        system_capacity_2 = quantity_2 * .400
    elif mod_watt == "SPR-X22-370-AC":
        system_capacity = quantity * .370
        system_capacity_2 = quantity_2 * .370
    elif mod_watt == "SPR-X22-360-AC":
        system_capacity = quantity * .360
        system_capacity_2 = quantity_2 * .360
    elif mod_watt == "SPR-X22-360":
        system_capacity = quantity * .360
        system_capacity_2 = quantity_2 * .360
    elif mod_watt == "SPR-X21-350-BLK-AC":
        system_capacity = quantity * .350
        system_capacity_2 = quantity_2 * .350
    elif mod_watt == "SPR-E20-327-AC":
        system_capacity = quantity * .327
        system_capacity_2 = quantity_2 * .327
    elif mod_watt == "SPR-E20-327":
        system_capacity = quantity * .327
        system_capacity_2 = quantity_2 * .327
    elif mod_watt == "SPR-E19-320-AC":
        system_capacity = quantity * .320
        system_capacity_2 = quantity_2 * .320
    else:
        pass

    system_capacity = str(system_capacity)
    system_capacity_2 = str(system_capacity_2)

    # *********************************************************************************************************

    # setting string variables for NREL/PVWATTS parameters ****************************************************

    address = address.replace(" ", "%20").strip()
    address = ("address=" + address + "&")
    new_tilt = ("tilt=" + new_tilt + "&")
    new_azimuth = ("azimuth=" + new_azimuth + "&")
    old_tilt = ("tilt=" + old_tilt + "&")
    old_azimuth = ("azimuth=" + old_azimuth + "&")
    losses_o = ("losses=" + losses_o + "&")
    losses_n = ("losses=" + losses_n + "&")
    module_type = ("module_type=" + module_type + "&")
    array_type = ("array_type=" + array_type + "&")
    system_capacity = ("system_capacity=" + system_capacity + "&")
    system_capacity_2 = ("system_capacity=" + system_capacity_2 + "&")

    api_param = "&api_key="
    old_query = address + old_tilt + old_azimuth + losses_o + module_type + array_type + system_capacity + api_param + creds.api_key
    new_query = address + new_tilt + new_azimuth + losses_n + module_type + array_type + system_capacity_2 + api_param + creds.api_key

    # parameters set for NREL/PVW API connection, preparing to make get call & parse data after 200 response *

    # preforming first requests to NREL/PVW API -- Looks a little jank but it works **************************
    response = requests.get("https://developer.nrel.gov/api/pvwatts/v6.json?" + api_param + creds.api_key)
    base_url = "https://developer.nrel.gov/api/pvwatts/v6.json?"

    json_link_original = (base_url + old_query)
    json_link_new = (base_url + new_query)
    data_original = requests.get(json_link_original)
    data_new = requests.get(json_link_new)

    print(json_link_original)
    print(json_link_new)
    print(data_original.status_code)
    print(data_new.status_code)

    # finished api request should have 200 response in console / printed *************************************

    # had a lot of trouble with the json formatting from NREL/PVW had to parse this way to get annual total **
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
        # Remove all characters after the character '.' from string
        ac_monthly_original = ac_monthly_original[0: ac_monthly_original.index(ch)]
    except ValueError:
        pass

    ch = '.'
    try:
        # Remove all characters after the character '.' from string
        ac_monthly_new = ac_monthly_new[0: ac_monthly_new.index(ch)]
    except ValueError:
        pass

    # ^^^^ finished parsing / quite the mess but works great, on to calculate percent difference between calls

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

    # calculations for ten percent docx sheet finished, some parsing of useless string data ******************

    # setting variables and values for ten percent docx and finishing out with a final print *****************
    # This is here to fix the TEN_PERCENT letter back to original formatting
    old_tilt = str(tab_lookup.acell("M16").value)
    old_azimuth = str(tab_lookup.acell("M17").value)
    new_tilt = str(tab_lookup.acell("N16").value)
    new_azimuth = str(tab_lookup.acell("N17").value)
    line_break = "________________________________________________"

    doc = DocxTemplate("TEN_PERCENT_V5.docx")
    context = {'hoa_name': hoa_name, 'date': date, 'name': name,
               'quantity': quantity, 'old_direction': old_direction, 'quantity2': quantity_2, 'state': state,
               'old_azimuth': old_azimuth, 'old_tilt': old_tilt, 'new_direction': new_direction,
               'new_azimuth': new_azimuth, 'new_tilt': new_tilt, 'mod_watt': mod_watt, 'percent': total,
               'ac_monthly_original': ac_monthly_original, 'ac_monthly_new': ac_monthly_new}

    doc.render(context)
    doc.save(name + " Ten Percent Letter array 3.docx")
    print("Ten Percent Letter finished...")

def main():

    if arrayCount == 1:
        arrayOne()
    elif arrayCount == 2:
        arrayOne()
        arrayTwo()
    elif arrayCount == 3:
        arrayOne()
        arrayTwo()
        arrayThree()
    elif arrayCount == 4:
        pass
    else:
        exit()

if __name__ == '__main__':
    main()