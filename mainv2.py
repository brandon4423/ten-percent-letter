from dataclasses import dataclass
import creds
import gspread
import requests
import time
import os
from docx2pdf import convert
from docxtpl import DocxTemplate
import multiprocessing as mp

login = gspread.service_account(filename="service_account.json")
sheet_name = login.open("HOA")
worksheet = sheet_name.worksheet("10 Percent")
values = worksheet.get_values("B1:F35")
location = values[12][0].replace(",,", ",")
date = values[21][0]

ch = '.'

class Customer:
    def __init__(self, name, state, hoa_name, address, mod_watt, array_count, module_type, array_type, pvwatts, pdf_response):
        self.name = name
        self.state = state
        self.hoa_name = hoa_name
        self.address = address
        self.mod_watt = mod_watt
        self.array_count = array_count
        self.module_type = module_type
        self.array_type = array_type
        self.pvwatts = pvwatts
        self.pdf_response = pdf_response
        
customer = Customer(values[12][3], values[3][2], values[6][0], values[12][0].replace(" ", "%20"), float(values[9][4]), int(values[17][1]), 1, 1, values[15][3].upper(), values[15][4].upper())

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
    print(query_1)

    all_data_1 = (response_1.json())
    all_data_2 = (response_2.json())

    json_data_1 = all_data_1.copy()
    json_data_2 = all_data_2.copy()

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
    if customer.pdf_response == 'YES':
        doc.save(customer.name + " Ten Percent Letter 1.docx")

        convert(f"{customer.name} Ten Percent Letter 1.docx", f"{customer.name} Ten Percent Letter 1.pdf")
        os.remove(f"{customer.name} Ten Percent Letter 1.docx")
    else:
        doc.save(customer.name + " Ten Percent Letter 1.docx")

    def pvcal():
        if customer.pvwatts == 'YES':
            monthly_data_1 = all_data_1.copy()
            monthly_data_2 = all_data_2.copy()
            monthly_data_1 = monthly_data_1['outputs']['ac_monthly']
            monthly_data_2 = monthly_data_2['outputs']['ac_monthly']

            annual_1 = all_data_1.copy()
            annual_2 = all_data_2.copy()
            annual_1 = all_data_1['outputs']['ac_annual']
            annual_2 = all_data_2['outputs']['ac_annual']
            
            monthly_sodata_1 = all_data_1.copy()
            monthly_sodata_2 = all_data_2.copy()
            monthly_sodata_1 = monthly_sodata_1['outputs']['solrad_monthly']
            monthly_sodata_2 = monthly_sodata_2['outputs']['solrad_monthly']

            annualso_1 = all_data_1.copy()
            annualso_2 = all_data_2.copy()
            annualso_1 = all_data_1['outputs']['solrad_annual']
            annualso_2 = all_data_2['outputs']['solrad_annual']

            lat = all_data_1.copy()
            lon = all_data_2.copy()
            lat = lat['station_info']['lat']
            lon = lon['station_info']['lon']

            capacity_total = all_data_1.copy()
            capacity_total = all_data_1['outputs']['capacity_factor']


            class Pvdata:
                def __init__(self, jan, feb, mar, apr, may, jun, jul, aug, sep, oct, nov, dec, annual, lat, lon, jan_so, feb_so, mar_so,
                apr_so, may_so, jun_so, jul_so, aug_so, sep_so, oct_so, nov_so, dec_so, annual_so, capacity_factor):
                    self.jan = jan
                    self.feb = feb
                    self.mar = mar
                    self.apr = apr
                    self.may = may
                    self.jun = jun
                    self.jul = jul
                    self.aug = aug
                    self.sep = sep
                    self.oct = oct
                    self.nov = nov
                    self.dec = dec
                    self.annual = annual
                    self.lat = lat
                    self.lon = lon
                    self.jan_so = jan_so
                    self.feb_so = feb_so
                    self.mar_so = mar_so
                    self.apr_so = apr_so
                    self.may_so = may_so
                    self.jun_so = jun_so
                    self.jul_so = jul_so
                    self.aug_so = aug_so
                    self.sep_so = sep_so
                    self.oct_so = oct_so
                    self.nov_so = nov_so
                    self.dec_so = dec_so
                    self.annual_so = annual_so
                    self.capacity_factor = capacity_factor
            
            monthly_1 = Pvdata(monthly_data_1[0], monthly_data_1[1], monthly_data_1[2], monthly_data_1[3], monthly_data_1[4], monthly_data_1[5], monthly_data_1[6], monthly_data_1[7], monthly_data_1[8]
            , monthly_data_1[9], monthly_data_1[10], monthly_data_1[11], annual_1, lat, lon, monthly_sodata_1[0], monthly_sodata_1[1], monthly_sodata_1[2], monthly_sodata_1[3], monthly_sodata_1[4]
            , monthly_sodata_1[5], monthly_sodata_1[6], monthly_sodata_1[7], monthly_sodata_1[8], monthly_sodata_1[9], monthly_sodata_1[10], monthly_sodata_1[11], annualso_1, capacity_total)

            monthly_2 = Pvdata(monthly_data_2[0], monthly_data_2[1], monthly_data_2[2], monthly_data_2[3], monthly_data_2[4], monthly_data_2[5], monthly_data_2[6], monthly_data_2[7], monthly_data_2[8]
            , monthly_data_2[9], monthly_data_2[10], monthly_data_2[11], annual_2, lat, lon, monthly_sodata_2[0], monthly_sodata_2[1], monthly_sodata_2[2], monthly_sodata_2[3], monthly_sodata_2[4]
            , monthly_sodata_2[5], monthly_sodata_2[6], monthly_sodata_2[7], monthly_sodata_2[8], monthly_sodata_2[9], monthly_sodata_2[10], monthly_sodata_2[11], annualso_2, capacity_total)

            monthly = [monthly_data_1[0], monthly_data_1[1], monthly_data_1[2], monthly_data_1[3], monthly_data_1[4], monthly_data_1[5], monthly_data_1[6], monthly_data_1[7], monthly_data_1[8]
            , monthly_data_1[9], monthly_data_1[10], monthly_data_1[11]
            , monthly_data_2[0], monthly_data_2[1], monthly_data_2[2], monthly_data_2[3], monthly_data_2[4], monthly_data_2[5], monthly_data_2[6], monthly_data_2[7], monthly_data_2[8]
            , monthly_data_2[9], monthly_data_2[10], monthly_data_2[11]]

            solar_rad = [monthly_1.jan_so, monthly_1.feb_so, monthly_1.mar_so, monthly_1.apr_so, monthly_1.may_so, monthly_1.jun_so
            , monthly_1.jul_so, monthly_1.aug_so, monthly_1.sep_so, monthly_1.oct_so, monthly_1.nov_so, monthly_1.dec_so,
            monthly_2.jan_so, monthly_2.feb_so, monthly_2.mar_so, monthly_2.apr_so, monthly_2.may_so, monthly_2.jun_so
            , monthly_2.jul_so, monthly_2.aug_so, monthly_2.sep_so, monthly_2.oct_so, monthly_2.nov_so, monthly_2.dec_so]

            solar_annual = [monthly_1.annual, monthly_2.annual]

            solar_annualrad = [monthly_1.annual_so, monthly_2.annual_so]

            location_data = [monthly_1.lat, monthly_1.lon, monthly_2.lat, monthly_2.lon]

            total_cap = [monthly_1.capacity_factor, monthly_2.capacity_factor]

            month = []
            for i in monthly:
                i = '{:0,.0f}'.format(i)
                month.append(i)

            rad = []
            for i in solar_rad:
                i = '{:0,.2f}'.format(i)
                rad.append(i)

            annual = []
            for i in solar_annual:
                i = '{:0,.0f}'.format(i)
                annual.append(i)

            annualrad = []
            for i in solar_annualrad:
                i = '{:0,.2f}'.format(i)
                annualrad.append(i)

            loc_data = []
            for i in location_data:
                i = '{:0,.2f}'.format(i)
                loc_data.append(i)
            
            capacity_factor = []
            for i in total_cap:
                i = '{:0,.1f}'.format(i)
                capacity_factor.append(i)

            solar_totals = [monthly_1.annual, monthly_2.annual]
            math_1 = solar_totals.copy()
            math_2 = solar_totals.copy()

            range1_percent = 0.024 * math_1[0]
            rangemath1 = math_1[0] - range1_percent

            range2_percent = 0.024 * math_1[1]
            rangemath2 = math_1[1] - range2_percent

            range3_percent = 0.02218 * math_2[0]
            rangemath3 = math_2[0] + range3_percent

            range4_percent = 0.02218 * math_2[1]
            rangemath4 = math_2[1] + range4_percent

            range_totals_1 = [rangemath1, rangemath2]
            range_totals_2 = [rangemath3, rangemath4]

            ranges_low = []
            for i in range_totals_1:
                i = '{:0,.0f}'.format(i)
                ranges_low.append(i)

            ranges_high = []
            for i in range_totals_2:
                i = '{:0,.0f}'.format(i)
                ranges_high.append(i)

            doc = DocxTemplate("LETTER_1.docx")
            context = {
                    'jan': month[0], 'jan_so': rad[0],
                    'feb': month[1], 'feb_so': rad[1],
                    'mar': month[2], 'mar_so': rad[2],
                    'apr': month[3], 'apr_so': rad[3],
                    'may': month[4], 'may_so': rad[4],
                    'jun': month[5], 'jun_so': rad[5],
                    'jul': month[6], 'jul_so': rad[6],
                    'aug': month[7], 'aug_so': rad[7],
                    'sep': month[8], 'sep_so': rad[8],
                    'oct': month[9], 'oct_so': rad[9],
                    'nov': month[10], 'nov_so': rad[10],
                    'dec': month[11], 'dec_so': rad[11],
                    'annual': annual[0], 'annual_so': annualrad[0],
                    'lat': loc_data[0], 'lon': loc_data[1], 'location': location,
                    'system_capacity': str(system_capacity_1) + ' kW',
                    'tilt': values[16][3], 'azimuth': values[17][3], 'losses': values[18][3],
                    'capacity_factor': capacity_factor[0], 'range1': ranges_low[0], 'range2': ranges_high[0]

                }

            doc.render(context)
            if customer.pdf_response == 'YES':
                doc.save(customer.name + " PVWatts Calculation 1.docx")

                convert(f"{customer.name} PVWatts Calculation 1.docx", f"{customer.name} PVWatts Calculation 1.pdf")
                os.remove(f"{customer.name} PVWatts Calculation 1.docx")
            else:
                doc.save(customer.name + " PVWatts Calculation 1.docx")

            doc = DocxTemplate("LETTER_2.docx")
            context = {
                    'jan': month[12], 'jan_so': rad[12],
                    'feb': month[13], 'feb_so': rad[13],
                    'mar': month[14], 'mar_so': rad[14],
                    'apr': month[15], 'apr_so': rad[15],
                    'may': month[16], 'may_so': rad[16],
                    'jun': month[17], 'jun_so': rad[17],
                    'jul': month[18], 'jul_so': rad[18],
                    'aug': month[19], 'aug_so': rad[19],
                    'sep': month[20], 'sep_so': rad[20],
                    'oct': month[21], 'oct_so': rad[21],
                    'nov': month[22], 'nov_so': rad[22],
                    'dec': month[23], 'dec_so': rad[23],
                    'annual': annual[1], 'annual_so': annualrad[1],
                    'lat': loc_data[2], 'lon': loc_data[3], 'location': location,
                    'system_capacity': str(system_capacity_1) + ' kW',
                    'tilt': values[16][3], 'azimuth': values[17][3], 'losses': values[18][4],
                    'capacity_factor': capacity_factor[1], 'range1': ranges_low[1], 'range2': ranges_high[1]

                }

            doc.render(context)
            if customer.pdf_response == 'YES':
                doc.save(customer.name + " PVWatts Calculation 2.docx")

                convert(f"{customer.name} PVWatts Calculation 2.docx", f"{customer.name} PVWatts Calculation 2.pdf")
                os.remove(f"{customer.name} PVWatts Calculation 2.docx")
            else:
                doc.save(customer.name + " PVWatts Calculation 2.docx")
        else:
            pass

    pvcal()
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

    all_data_1 = (response_3.json())
    all_data_2 = (response_4.json())

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
    if customer.pdf_response == 'YES':
        doc.save(customer.name + " Ten Percent Letter 2.docx")

        convert(f"{customer.name} Ten Percent Letter 2.docx", f"{customer.name} Ten Percent Letter 2.pdf")
        os.remove(f"{customer.name} Ten Percent Letter 2.docx")
    else:
        doc.save(customer.name + " Ten Percent Letter 2.docx")

    def pvcal():
        if customer.pvwatts == 'YES':
            monthly_data_1 = all_data_1.copy()
            monthly_data_2 = all_data_2.copy()
            monthly_data_1 = monthly_data_1['outputs']['ac_monthly']
            monthly_data_2 = monthly_data_2['outputs']['ac_monthly']

            annual_1 = all_data_1.copy()
            annual_2 = all_data_2.copy()
            annual_1 = all_data_1['outputs']['ac_annual']
            annual_2 = all_data_2['outputs']['ac_annual']
            
            monthly_sodata_1 = all_data_1.copy()
            monthly_sodata_2 = all_data_2.copy()
            monthly_sodata_1 = monthly_sodata_1['outputs']['solrad_monthly']
            monthly_sodata_2 = monthly_sodata_2['outputs']['solrad_monthly']

            annualso_1 = all_data_1.copy()
            annualso_2 = all_data_2.copy()
            annualso_1 = all_data_1['outputs']['solrad_annual']
            annualso_2 = all_data_2['outputs']['solrad_annual']

            lat = all_data_1.copy()
            lon = all_data_2.copy()
            lat = lat['station_info']['lat']
            lon = lon['station_info']['lon']

            capacity_total = all_data_1.copy()
            capacity_total = all_data_1['outputs']['capacity_factor']


            class Pvdata:
                def __init__(self, jan, feb, mar, apr, may, jun, jul, aug, sep, oct, nov, dec, annual, lat, lon, jan_so, feb_so, mar_so,
                apr_so, may_so, jun_so, jul_so, aug_so, sep_so, oct_so, nov_so, dec_so, annual_so, capacity_factor):
                    self.jan = jan
                    self.feb = feb
                    self.mar = mar
                    self.apr = apr
                    self.may = may
                    self.jun = jun
                    self.jul = jul
                    self.aug = aug
                    self.sep = sep
                    self.oct = oct
                    self.nov = nov
                    self.dec = dec
                    self.annual = annual
                    self.lat = lat
                    self.lon = lon
                    self.jan_so = jan_so
                    self.feb_so = feb_so
                    self.mar_so = mar_so
                    self.apr_so = apr_so
                    self.may_so = may_so
                    self.jun_so = jun_so
                    self.jul_so = jul_so
                    self.aug_so = aug_so
                    self.sep_so = sep_so
                    self.oct_so = oct_so
                    self.nov_so = nov_so
                    self.dec_so = dec_so
                    self.annual_so = annual_so
                    self.capacity_factor = capacity_factor
            
            monthly_1 = Pvdata(monthly_data_1[0], monthly_data_1[1], monthly_data_1[2], monthly_data_1[3], monthly_data_1[4], monthly_data_1[5], monthly_data_1[6], monthly_data_1[7], monthly_data_1[8]
            , monthly_data_1[9], monthly_data_1[10], monthly_data_1[11], annual_1, lat, lon, monthly_sodata_1[0], monthly_sodata_1[1], monthly_sodata_1[2], monthly_sodata_1[3], monthly_sodata_1[4]
            , monthly_sodata_1[5], monthly_sodata_1[6], monthly_sodata_1[7], monthly_sodata_1[8], monthly_sodata_1[9], monthly_sodata_1[10], monthly_sodata_1[11], annualso_1, capacity_total)

            monthly_2 = Pvdata(monthly_data_2[0], monthly_data_2[1], monthly_data_2[2], monthly_data_2[3], monthly_data_2[4], monthly_data_2[5], monthly_data_2[6], monthly_data_2[7], monthly_data_2[8]
            , monthly_data_2[9], monthly_data_2[10], monthly_data_2[11], annual_2, lat, lon, monthly_sodata_2[0], monthly_sodata_2[1], monthly_sodata_2[2], monthly_sodata_2[3], monthly_sodata_2[4]
            , monthly_sodata_2[5], monthly_sodata_2[6], monthly_sodata_2[7], monthly_sodata_2[8], monthly_sodata_2[9], monthly_sodata_2[10], monthly_sodata_2[11], annualso_2, capacity_total)

            monthly = [monthly_data_1[0], monthly_data_1[1], monthly_data_1[2], monthly_data_1[3], monthly_data_1[4], monthly_data_1[5], monthly_data_1[6], monthly_data_1[7], monthly_data_1[8]
            , monthly_data_1[9], monthly_data_1[10], monthly_data_1[11]
            , monthly_data_2[0], monthly_data_2[1], monthly_data_2[2], monthly_data_2[3], monthly_data_2[4], monthly_data_2[5], monthly_data_2[6], monthly_data_2[7], monthly_data_2[8]
            , monthly_data_2[9], monthly_data_2[10], monthly_data_2[11]]

            solar_rad = [monthly_1.jan_so, monthly_1.feb_so, monthly_1.mar_so, monthly_1.apr_so, monthly_1.may_so, monthly_1.jun_so
            , monthly_1.jul_so, monthly_1.aug_so, monthly_1.sep_so, monthly_1.oct_so, monthly_1.nov_so, monthly_1.dec_so,
            monthly_2.jan_so, monthly_2.feb_so, monthly_2.mar_so, monthly_2.apr_so, monthly_2.may_so, monthly_2.jun_so
            , monthly_2.jul_so, monthly_2.aug_so, monthly_2.sep_so, monthly_2.oct_so, monthly_2.nov_so, monthly_2.dec_so]

            solar_annual = [monthly_1.annual, monthly_2.annual]

            solar_annualrad = [monthly_1.annual_so, monthly_2.annual_so]

            location_data = [monthly_1.lat, monthly_1.lon, monthly_2.lat, monthly_2.lon]

            total_cap = [monthly_1.capacity_factor, monthly_2.capacity_factor]

            month = []
            for i in monthly:
                i = '{:0,.0f}'.format(i)
                month.append(i)

            rad = []
            for i in solar_rad:
                i = '{:0,.2f}'.format(i)
                rad.append(i)

            annual = []
            for i in solar_annual:
                i = '{:0,.0f}'.format(i)
                annual.append(i)

            annualrad = []
            for i in solar_annualrad:
                i = '{:0,.2f}'.format(i)
                annualrad.append(i)

            loc_data = []
            for i in location_data:
                i = '{:0,.2f}'.format(i)
                loc_data.append(i)
            
            capacity_factor = []
            for i in total_cap:
                i = '{:0,.1f}'.format(i)
                capacity_factor.append(i)

            solar_totals = [monthly_1.annual, monthly_2.annual]
            math_1 = solar_totals.copy()
            math_2 = solar_totals.copy()

            range1_percent = 0.024 * math_1[0]
            rangemath1 = math_1[0] - range1_percent

            range2_percent = 0.024 * math_1[1]
            rangemath2 = math_1[1] - range2_percent

            range3_percent = 0.02218 * math_2[0]
            rangemath3 = math_2[0] + range3_percent

            range4_percent = 0.02218 * math_2[1]
            rangemath4 = math_2[1] + range4_percent

            range_totals_1 = [rangemath1, rangemath2]
            range_totals_2 = [rangemath3, rangemath4]

            ranges_low = []
            for i in range_totals_1:
                i = '{:0,.0f}'.format(i)
                ranges_low.append(i)

            ranges_high = []
            for i in range_totals_2:
                i = '{:0,.0f}'.format(i)
                ranges_high.append(i)

            doc = DocxTemplate("LETTER_1.docx")
            context = {
                    'jan': month[0], 'jan_so': rad[0],
                    'feb': month[1], 'feb_so': rad[1],
                    'mar': month[2], 'mar_so': rad[2],
                    'apr': month[3], 'apr_so': rad[3],
                    'may': month[4], 'may_so': rad[4],
                    'jun': month[5], 'jun_so': rad[5],
                    'jul': month[6], 'jul_so': rad[6],
                    'aug': month[7], 'aug_so': rad[7],
                    'sep': month[8], 'sep_so': rad[8],
                    'oct': month[9], 'oct_so': rad[9],
                    'nov': month[10], 'nov_so': rad[10],
                    'dec': month[11], 'dec_so': rad[11],
                    'annual': annual[0], 'annual_so': annualrad[0],
                    'lat': loc_data[0], 'lon': loc_data[1], 'location': location,
                    'system_capacity': str(system_capacity_3) + ' kW',
                    'tilt': values[23][3], 'azimuth': values[24][3], 'losses': values[25][3],
                    'capacity_factor': capacity_factor[0], 'range1': ranges_low[0], 'range2': ranges_high[0]

                }

            doc.render(context)
            if customer.pdf_response == 'YES':
                doc.save(customer.name + " PVWatts Calculation 3.docx")

                convert(f"{customer.name} PVWatts Calculation 3.docx", f"{customer.name} PVWatts Calculation 3.pdf")
                os.remove(f"{customer.name} PVWatts Calculation 3.docx")
            else:
                doc.save(customer.name + " PVWatts Calculation 3.docx")

            doc = DocxTemplate("LETTER_2.docx")
            context = {
                    'jan': month[12], 'jan_so': rad[12],
                    'feb': month[13], 'feb_so': rad[13],
                    'mar': month[14], 'mar_so': rad[14],
                    'apr': month[15], 'apr_so': rad[15],
                    'may': month[16], 'may_so': rad[16],
                    'jun': month[17], 'jun_so': rad[17],
                    'jul': month[18], 'jul_so': rad[18],
                    'aug': month[19], 'aug_so': rad[19],
                    'sep': month[20], 'sep_so': rad[20],
                    'oct': month[21], 'oct_so': rad[21],
                    'nov': month[22], 'nov_so': rad[22],
                    'dec': month[23], 'dec_so': rad[23],
                    'annual': annual[1], 'annual_so': annualrad[1],
                    'lat': loc_data[2], 'lon': loc_data[3], 'location': location,
                    'system_capacity': str(system_capacity_4) + ' kW',
                    'tilt': values[23][4], 'azimuth': values[24][4], 'losses': values[25][4],
                    'capacity_factor': capacity_factor[1], 'range1': ranges_low[1], 'range2': ranges_high[1]

                }

            doc.render(context)
            if customer.pdf_response == 'YES':
                doc.save(customer.name + " PVWatts Calculation 4.docx")

                convert(f"{customer.name} PVWatts Calculation 4.docx", f"{customer.name} PVWatts Calculation 4.pdf")
                os.remove(f"{customer.name} PVWatts Calculation 4.docx")
            else:
                doc.save(customer.name + " PVWatts Calculation 4.docx")
        else:
            pass

    pvcal()
    end_time = time.time()
    final_time = end_time - start_time
    print(f"Letter One run in: {final_time} seconds")

def letterThree():
    start_time = time.time()
    system_capacity_5 = customer.mod_watt * array_5.quantity
    system_capacity_6 = customer.mod_watt * array_6.quantity

    query_5 = (f"&api_key={creds.api_key}&address={customer.address}&tilt={array_5.tilt}&azimuth={array_5.azimuth}&losses={array_5.losses}&module_type={customer.module_type}&array_type={customer.array_type}&system_capacity={system_capacity_5}")
    query_6 = (f"&api_key={creds.api_key}&address={customer.address}&tilt={array_6.tilt}&azimuth={array_6.azimuth}&losses={array_6.losses}&module_type={customer.module_type}&array_type={customer.array_type}&system_capacity={system_capacity_6}")

    response_5 = requests.get("https://developer.nrel.gov/api/pvwatts/v6.json?" + query_5)
    response_6 = requests.get("https://developer.nrel.gov/api/pvwatts/v6.json?" + query_6)

    all_data_1 = (response_5.json())
    all_data_2 = (response_6.json())

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
    if customer.pdf_response == 'YES':
        doc.save(customer.name + " Ten Percent Letter 3.docx")

        convert(f"{customer.name} Ten Percent Letter 3.docx", f"{customer.name} Ten Percent Letter 3.pdf")
        os.remove(f"{customer.name} Ten Percent Letter 3.docx")
    else:
        doc.save(customer.name + " Ten Percent Letter 3.docx")

    def pvcal():
        if customer.pvwatts == 'YES':
            monthly_data_1 = all_data_1.copy()
            monthly_data_2 = all_data_2.copy()
            monthly_data_1 = monthly_data_1['outputs']['ac_monthly']
            monthly_data_2 = monthly_data_2['outputs']['ac_monthly']

            annual_1 = all_data_1.copy()
            annual_2 = all_data_2.copy()
            annual_1 = all_data_1['outputs']['ac_annual']
            annual_2 = all_data_2['outputs']['ac_annual']
            
            monthly_sodata_1 = all_data_1.copy()
            monthly_sodata_2 = all_data_2.copy()
            monthly_sodata_1 = monthly_sodata_1['outputs']['solrad_monthly']
            monthly_sodata_2 = monthly_sodata_2['outputs']['solrad_monthly']

            annualso_1 = all_data_1.copy()
            annualso_2 = all_data_2.copy()
            annualso_1 = all_data_1['outputs']['solrad_annual']
            annualso_2 = all_data_2['outputs']['solrad_annual']

            lat = all_data_1.copy()
            lon = all_data_2.copy()
            lat = lat['station_info']['lat']
            lon = lon['station_info']['lon']

            capacity_total = all_data_1.copy()
            capacity_total = all_data_1['outputs']['capacity_factor']


            class Pvdata:
                def __init__(self, jan, feb, mar, apr, may, jun, jul, aug, sep, oct, nov, dec, annual, lat, lon, jan_so, feb_so, mar_so,
                apr_so, may_so, jun_so, jul_so, aug_so, sep_so, oct_so, nov_so, dec_so, annual_so, capacity_factor):
                    self.jan = jan
                    self.feb = feb
                    self.mar = mar
                    self.apr = apr
                    self.may = may
                    self.jun = jun
                    self.jul = jul
                    self.aug = aug
                    self.sep = sep
                    self.oct = oct
                    self.nov = nov
                    self.dec = dec
                    self.annual = annual
                    self.lat = lat
                    self.lon = lon
                    self.jan_so = jan_so
                    self.feb_so = feb_so
                    self.mar_so = mar_so
                    self.apr_so = apr_so
                    self.may_so = may_so
                    self.jun_so = jun_so
                    self.jul_so = jul_so
                    self.aug_so = aug_so
                    self.sep_so = sep_so
                    self.oct_so = oct_so
                    self.nov_so = nov_so
                    self.dec_so = dec_so
                    self.annual_so = annual_so
                    self.capacity_factor = capacity_factor
            
            monthly_1 = Pvdata(monthly_data_1[0], monthly_data_1[1], monthly_data_1[2], monthly_data_1[3], monthly_data_1[4], monthly_data_1[5], monthly_data_1[6], monthly_data_1[7], monthly_data_1[8]
            , monthly_data_1[9], monthly_data_1[10], monthly_data_1[11], annual_1, lat, lon, monthly_sodata_1[0], monthly_sodata_1[1], monthly_sodata_1[2], monthly_sodata_1[3], monthly_sodata_1[4]
            , monthly_sodata_1[5], monthly_sodata_1[6], monthly_sodata_1[7], monthly_sodata_1[8], monthly_sodata_1[9], monthly_sodata_1[10], monthly_sodata_1[11], annualso_1, capacity_total)

            monthly_2 = Pvdata(monthly_data_2[0], monthly_data_2[1], monthly_data_2[2], monthly_data_2[3], monthly_data_2[4], monthly_data_2[5], monthly_data_2[6], monthly_data_2[7], monthly_data_2[8]
            , monthly_data_2[9], monthly_data_2[10], monthly_data_2[11], annual_2, lat, lon, monthly_sodata_2[0], monthly_sodata_2[1], monthly_sodata_2[2], monthly_sodata_2[3], monthly_sodata_2[4]
            , monthly_sodata_2[5], monthly_sodata_2[6], monthly_sodata_2[7], monthly_sodata_2[8], monthly_sodata_2[9], monthly_sodata_2[10], monthly_sodata_2[11], annualso_2, capacity_total)

            monthly = [monthly_data_1[0], monthly_data_1[1], monthly_data_1[2], monthly_data_1[3], monthly_data_1[4], monthly_data_1[5], monthly_data_1[6], monthly_data_1[7], monthly_data_1[8]
            , monthly_data_1[9], monthly_data_1[10], monthly_data_1[11]
            , monthly_data_2[0], monthly_data_2[1], monthly_data_2[2], monthly_data_2[3], monthly_data_2[4], monthly_data_2[5], monthly_data_2[6], monthly_data_2[7], monthly_data_2[8]
            , monthly_data_2[9], monthly_data_2[10], monthly_data_2[11]]

            solar_rad = [monthly_1.jan_so, monthly_1.feb_so, monthly_1.mar_so, monthly_1.apr_so, monthly_1.may_so, monthly_1.jun_so
            , monthly_1.jul_so, monthly_1.aug_so, monthly_1.sep_so, monthly_1.oct_so, monthly_1.nov_so, monthly_1.dec_so,
            monthly_2.jan_so, monthly_2.feb_so, monthly_2.mar_so, monthly_2.apr_so, monthly_2.may_so, monthly_2.jun_so
            , monthly_2.jul_so, monthly_2.aug_so, monthly_2.sep_so, monthly_2.oct_so, monthly_2.nov_so, monthly_2.dec_so]

            solar_annual = [monthly_1.annual, monthly_2.annual]

            solar_annualrad = [monthly_1.annual_so, monthly_2.annual_so]

            location_data = [monthly_1.lat, monthly_1.lon, monthly_2.lat, monthly_2.lon]

            total_cap = [monthly_1.capacity_factor, monthly_2.capacity_factor]

            month = []
            for i in monthly:
                i = '{:0,.0f}'.format(i)
                month.append(i)

            rad = []
            for i in solar_rad:
                i = '{:0,.2f}'.format(i)
                rad.append(i)

            annual = []
            for i in solar_annual:
                i = '{:0,.0f}'.format(i)
                annual.append(i)

            annualrad = []
            for i in solar_annualrad:
                i = '{:0,.2f}'.format(i)
                annualrad.append(i)

            loc_data = []
            for i in location_data:
                i = '{:0,.2f}'.format(i)
                loc_data.append(i)
            
            capacity_factor = []
            for i in total_cap:
                i = '{:0,.1f}'.format(i)
                capacity_factor.append(i)

            solar_totals = [monthly_1.annual, monthly_2.annual]
            math_1 = solar_totals.copy()
            math_2 = solar_totals.copy()

            range1_percent = 0.024 * math_1[0]
            rangemath1 = math_1[0] - range1_percent

            range2_percent = 0.024 * math_1[1]
            rangemath2 = math_1[1] - range2_percent

            range3_percent = 0.02218 * math_2[0]
            rangemath3 = math_2[0] + range3_percent

            range4_percent = 0.02218 * math_2[1]
            rangemath4 = math_2[1] + range4_percent

            range_totals_1 = [rangemath1, rangemath2]
            range_totals_2 = [rangemath3, rangemath4]

            ranges_low = []
            for i in range_totals_1:
                i = '{:0,.0f}'.format(i)
                ranges_low.append(i)

            ranges_high = []
            for i in range_totals_2:
                i = '{:0,.0f}'.format(i)
                ranges_high.append(i)

            doc = DocxTemplate("LETTER_1.docx")
            context = {
                    'jan': month[0], 'jan_so': rad[0],
                    'feb': month[1], 'feb_so': rad[1],
                    'mar': month[2], 'mar_so': rad[2],
                    'apr': month[3], 'apr_so': rad[3],
                    'may': month[4], 'may_so': rad[4],
                    'jun': month[5], 'jun_so': rad[5],
                    'jul': month[6], 'jul_so': rad[6],
                    'aug': month[7], 'aug_so': rad[7],
                    'sep': month[8], 'sep_so': rad[8],
                    'oct': month[9], 'oct_so': rad[9],
                    'nov': month[10], 'nov_so': rad[10],
                    'dec': month[11], 'dec_so': rad[11],
                    'annual': annual[0], 'annual_so': annualrad[0],
                    'lat': loc_data[0], 'lon': loc_data[1], 'location': location,
                    'system_capacity': str(system_capacity_5) + ' kW',
                    'tilt': values[30][3], 'azimuth': values[31][3], 'losses': values[32][3],
                    'capacity_factor': capacity_factor[0], 'range1': ranges_low[0], 'range2': ranges_high[0]

                }

            doc.render(context)
            if customer.pdf_response == 'YES':
                doc.save(customer.name + " PVWatts Calculation 5.docx")

                convert(f"{customer.name} PVWatts Calculation 5.docx", f"{customer.name} PVWatts Calculation 5.pdf")
                os.remove(f"{customer.name} PVWatts Calculation 5.docx")
            else:
                doc.save(customer.name + " PVWatts Calculation 5.docx")

            doc = DocxTemplate("LETTER_2.docx")
            context = {
                    'jan': month[12], 'jan_so': rad[12],
                    'feb': month[13], 'feb_so': rad[13],
                    'mar': month[14], 'mar_so': rad[14],
                    'apr': month[15], 'apr_so': rad[15],
                    'may': month[16], 'may_so': rad[16],
                    'jun': month[17], 'jun_so': rad[17],
                    'jul': month[18], 'jul_so': rad[18],
                    'aug': month[19], 'aug_so': rad[19],
                    'sep': month[20], 'sep_so': rad[20],
                    'oct': month[21], 'oct_so': rad[21],
                    'nov': month[22], 'nov_so': rad[22],
                    'dec': month[23], 'dec_so': rad[23],
                    'annual': annual[1], 'annual_so': annualrad[1],
                    'lat': loc_data[2], 'lon': loc_data[3], 'location': location,
                    'system_capacity': str(system_capacity_6) + ' kW',
                    'tilt': values[30][4], 'azimuth': values[31][4], 'losses': values[32][4],
                    'capacity_factor': capacity_factor[1], 'range1': ranges_low[1], 'range2': ranges_high[1]

                }

            doc.render(context)
            if customer.pdf_response == 'YES':
                doc.save(customer.name + " PVWatts Calculation 6.docx")

                convert(f"{customer.name} PVWatts Calculation 6.docx", f"{customer.name} PVWatts Calculation 6.pdf")
                os.remove(f"{customer.name} PVWatts Calculation 6.docx")
            else:
                doc.save(customer.name + " PVWatts Calculation 6.docx")
        else:
            pass

    pvcal()
    end_time = time.time()
    final_time = end_time - start_time
    print(f"Letter One run in: {final_time} seconds")

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