import os
import time
import creds
import tenref
import gspread
import requests
from docx2pdf import convert
from dataclasses import dataclass
from docxtpl import DocxTemplate

login = gspread.service_account(filename="service_account.json")
sheet_name = login.open("HOA")
worksheet = sheet_name.worksheet("10 Percent")
values = worksheet.get_values("B1:F35")
location_address = values[12][0].replace(",,", ",")
date = values[21][0]

ch = '.'

@dataclass
class Customer:
        name: str
        state: str
        hoa_name: str
        address: str
        mod_watt: float
        array_count: int
        module_type: int
        array_type: int
        pvwatts: str
        pdf_response: str
        
customer = Customer(values[12][3], values[3][2], values[6][0], values[12][0].replace(" ", "%20"), float(values[9][4]), int(values[17][1]), 1, 1, values[15][3].upper(), values[15][4].upper())

@dataclass
class Array:
        tilt: str
        azimuth: str
        losses: str
        quantity: str
        direction: str

array_1 = Array(values[16][3], values[17][3], values[18][3], int(values[19][3]), values[20][3])
array_2 = Array(values[16][4], values[17][4], values[18][4], int(values[19][4]), values[20][4])
array_3 = Array(values[23][3], values[24][3], values[25][3], int(values[26][3]), values[27][3])
array_4 = Array(values[23][4], values[24][4], values[25][4], int(values[26][4]), values[27][4])
array_5 = Array(values[30][3], values[31][3], values[32][3], int(values[33][3]), values[34][3])
array_6 = Array(values[30][4], values[31][4], values[32][4], int(values[33][4]), values[34][4])

system_capacity_1 = customer.mod_watt * array_1.quantity
system_capacity_2 = customer.mod_watt * array_2.quantity

if customer.state == 'TX':
    customer.state = tenref.texas
elif customer.state == 'CO':
    customer.state = tenref.colorado
else:
    pass

def letterOne():
    start_time = time.time()

    query_1 = (f"&api_key={creds.api_key}&address={customer.address}&tilt={array_1.tilt}&azimuth={array_1.azimuth}&losses={array_1.losses}&module_type={customer.module_type}&array_type={customer.array_type}&system_capacity={system_capacity_1}")
    query_2 = (f"&api_key={creds.api_key}&address={customer.address}&tilt={array_2.tilt}&azimuth={array_2.azimuth}&losses={array_2.losses}&module_type={customer.module_type}&array_type={customer.array_type}&system_capacity={system_capacity_2}")

    response_1 = requests.get("https://developer.nrel.gov/api/pvwatts/v6.json?" + query_1)
    response_2 = requests.get("https://developer.nrel.gov/api/pvwatts/v6.json?" + query_2)

    data1 = (response_1.json())
    data2 = (response_2.json())

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
    if customer.pdf_response == 'YES':
        doc.save(customer.name + " Ten Percent Letter 1.docx")

        convert(f"{customer.name} Ten Percent Letter 1.docx", f"{customer.name} Ten Percent Letter 1.pdf")
        os.remove(f"{customer.name} Ten Percent Letter 1.docx")
    else:
        doc.save(customer.name + " Ten Percent Letter 1.docx")

    def pvcal():
        if customer.pvwatts == 'YES':
            jsonmonthly1 = data1['outputs']['ac_monthly']
            jsonmonthly2 = data2['outputs']['ac_monthly']
            jsonmonthlyrad1 = data1['outputs']['solrad_monthly']
            jsonmonthlyrad2 = data2['outputs']['solrad_monthly']
            jsonannual1 = data1['outputs']['ac_annual']
            jsonannual2 = data2['outputs']['ac_annual']
            jsonannualrad1 = data1['outputs']['solrad_annual']
            jsonannualrad2 = data2['outputs']['solrad_annual']
            lat = data1['station_info']['lat']
            lon = data2['station_info']['lon']
            jsoncap1 = data1['outputs']['capacity_factor']
            jsoncap2 = data2['outputs']['capacity_factor']
            
            rawannual = [jsonannual1, jsonannual2]
            rawannualrad = [jsonannualrad1, jsonannualrad2] 
            rawcap = [jsoncap1, jsoncap2]
            rawlocation = [lat, lon]

            monthly = []
            for x in jsonmonthly1:
                x = '{:0,.0f}'.format(x)
                monthly.append(x)
            for x in jsonmonthly2:
                x = '{:0,.0f}'.format(x)
                monthly.append(x)

            capacity = []
            for x in rawcap:
                x = '{:0,.1f}'.format(x)
                capacity.append(x)

            annualrad = []
            monthlyrad = []
            location = []
            for x in jsonmonthlyrad1:
                x = '{:0,.2f}'.format(x)
                monthlyrad.append(x)
            for x in jsonmonthlyrad2:
                x = '{:0,.2f}'.format(x)
                monthlyrad.append(x)
            for x in rawannualrad:
                x = '{:0,.2f}'.format(x)
                annualrad.append(x)
            for x in rawlocation:
                x = '{:0,.2f}'.format(x)
                location.append(x)

            solar_totals = [rawannual[0], rawannual[1]]
            low = [solar_totals[0]*0.024, solar_totals[1]*0.024]
            high = [solar_totals[0]*0.02218, solar_totals[1]*0.02218]
            low = [solar_totals[0] - low[0], solar_totals[1] - low[1]]
            high = [solar_totals[0] + high[0], solar_totals[1] + high[1]]

            ranges_low = []
            ranges_high = []
            annual = []
            for i in low:
                i = '{:0,.0f}'.format(i)
                ranges_low.append(i)
            for i in high:
                i = '{:0,.0f}'.format(i)
                ranges_high.append(i)
            for x in rawannual:
                x = '{:0,.0f}'.format(x)
                annual.append(x)
                    
            doc = DocxTemplate("LETTER.docx")
            context = {
                    'jan': monthly[0], 'jan_so': monthlyrad[0],
                    'feb': monthly[1], 'feb_so': monthlyrad[1],
                    'mar': monthly[2], 'mar_so': monthlyrad[2],
                    'apr': monthly[3], 'apr_so': monthlyrad[3],
                    'may': monthly[4], 'may_so': monthlyrad[4],
                    'jun': monthly[5], 'jun_so': monthlyrad[5],
                    'jul': monthly[6], 'jul_so': monthlyrad[6],
                    'aug': monthly[7], 'aug_so': monthlyrad[7],
                    'sep': monthly[8], 'sep_so': monthlyrad[8],
                    'oct': monthly[9], 'oct_so': monthlyrad[9],
                    'nov': monthly[10], 'nov_so': monthlyrad[10],
                    'dec': monthly[11], 'dec_so': monthlyrad[11],
                    'annual': annual[0], 'annual_so': annualrad[0],
                    'lat': location[0], 'lon': location[1], 'location': location_address,
                    'system_capacity': str(system_capacity_1) + ' kW',
                    'tilt': values[16][3], 'azimuth': values[17][3], 'losses': values[18][3],
                    'capacity_factor': capacity[0], 'range1': ranges_low[0], 'range2': ranges_high[0]

                }

            doc.render(context)
            if customer.pdf_response == 'YES':
                doc.save(customer.name + " PVWatts Calculation 1.docx")

                convert(f"{customer.name} PVWatts Calculation 1.docx", f"{customer.name} PVWatts Calculation 1.pdf")
                os.remove(f"{customer.name} PVWatts Calculation 1.docx")
            else:
                doc.save(customer.name + " PVWatts Calculation 1.docx")

            doc = DocxTemplate("LETTER.docx")
            context = {
                    'jan': monthly[12], 'jan_so': monthlyrad[12],
                    'feb': monthly[13], 'feb_so': monthlyrad[13],
                    'mar': monthly[14], 'mar_so': monthlyrad[14],
                    'apr': monthly[15], 'apr_so': monthlyrad[15],
                    'may': monthly[16], 'may_so': monthlyrad[16],
                    'jun': monthly[17], 'jun_so': monthlyrad[17],
                    'jul': monthly[18], 'jul_so': monthlyrad[18],
                    'aug': monthly[19], 'aug_so': monthlyrad[19],
                    'sep': monthly[20], 'sep_so': monthlyrad[20],
                    'oct': monthly[21], 'oct_so': monthlyrad[21],
                    'nov': monthly[22], 'nov_so': monthlyrad[22],
                    'dec': monthly[23], 'dec_so': monthlyrad[23],
                    'annual': annual[1], 'annual_so': annualrad[1],
                    'lat': location[0], 'lon': location[1], 'location': location_address,
                    'system_capacity': str(system_capacity_1) + ' kW',
                    'tilt': values[16][4], 'azimuth': values[17][4], 'losses': values[18][4],
                    'capacity_factor': capacity[1], 'range1': ranges_low[1], 'range2': ranges_high[1]

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

    data1 = (response_3.json())
    data2 = (response_4.json())

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
            jsonmonthly1 = data1['outputs']['ac_monthly']
            jsonmonthly2 = data2['outputs']['ac_monthly']
            jsonmonthlyrad1 = data1['outputs']['solrad_monthly']
            jsonmonthlyrad2 = data2['outputs']['solrad_monthly']
            jsonannual1 = data1['outputs']['ac_annual']
            jsonannual2 = data2['outputs']['ac_annual']
            jsonannualrad1 = data1['outputs']['solrad_annual']
            jsonannualrad2 = data2['outputs']['solrad_annual']
            lat = data1['station_info']['lat']
            lon = data2['station_info']['lon']
            jsoncap1 = data1['outputs']['capacity_factor']
            jsoncap2 = data2['outputs']['capacity_factor']
            
            rawannual = [jsonannual1, jsonannual2]
            rawannualrad = [jsonannualrad1, jsonannualrad2] 
            rawcap = [jsoncap1, jsoncap2]
            rawlocation = [lat, lon]

            monthly = []
            for x in jsonmonthly1:
                x = '{:0,.0f}'.format(x)
                monthly.append(x)
            for x in jsonmonthly2:
                x = '{:0,.0f}'.format(x)
                monthly.append(x)

            capacity = []
            for x in rawcap:
                x = '{:0,.1f}'.format(x)
                capacity.append(x)

            annualrad = []
            monthlyrad = []
            location = []
            for x in jsonmonthlyrad1:
                x = '{:0,.2f}'.format(x)
                monthlyrad.append(x)
            for x in jsonmonthlyrad2:
                x = '{:0,.2f}'.format(x)
                monthlyrad.append(x)
            for x in rawannualrad:
                x = '{:0,.2f}'.format(x)
                annualrad.append(x)
            for x in rawlocation:
                x = '{:0,.2f}'.format(x)
                location.append(x)

            solar_totals = [rawannual[0], rawannual[1]]
            low = [solar_totals[0]*0.024, solar_totals[1]*0.024]
            high = [solar_totals[0]*0.02218, solar_totals[1]*0.02218]
            low = [solar_totals[0] - low[0], solar_totals[1] - low[1]]
            high = [solar_totals[0] + high[0], solar_totals[1] + high[1]]

            ranges_low = []
            ranges_high = []
            annual = []
            for i in low:
                i = '{:0,.0f}'.format(i)
                ranges_low.append(i)
            for i in high:
                i = '{:0,.0f}'.format(i)
                ranges_high.append(i)
            for x in rawannual:
                x = '{:0,.0f}'.format(x)
                annual.append(x)
                    
            doc = DocxTemplate("LETTER.docx")
            context = {
                    'jan': monthly[0], 'jan_so': monthlyrad[0],
                    'feb': monthly[1], 'feb_so': monthlyrad[1],
                    'mar': monthly[2], 'mar_so': monthlyrad[2],
                    'apr': monthly[3], 'apr_so': monthlyrad[3],
                    'may': monthly[4], 'may_so': monthlyrad[4],
                    'jun': monthly[5], 'jun_so': monthlyrad[5],
                    'jul': monthly[6], 'jul_so': monthlyrad[6],
                    'aug': monthly[7], 'aug_so': monthlyrad[7],
                    'sep': monthly[8], 'sep_so': monthlyrad[8],
                    'oct': monthly[9], 'oct_so': monthlyrad[9],
                    'nov': monthly[10], 'nov_so': monthlyrad[10],
                    'dec': monthly[11], 'dec_so': monthlyrad[11],
                    'annual': annual[0], 'annual_so': annualrad[0],
                    'lat': location[0], 'lon': location[1], 'location': location_address,
                    'system_capacity': str(system_capacity_1) + ' kW',
                    'tilt': values[23][3], 'azimuth': values[24][3], 'losses': values[25][3],
                    'capacity_factor': capacity[0], 'range1': ranges_low[0], 'range2': ranges_high[0]

                }

            doc.render(context)
            if customer.pdf_response == 'YES':
                doc.save(customer.name + " PVWatts Calculation 3.docx")

                convert(f"{customer.name} PVWatts Calculation 3.docx", f"{customer.name} PVWatts Calculation 3.pdf")
                os.remove(f"{customer.name} PVWatts Calculation 3.docx")
            else:
                doc.save(customer.name + " PVWatts Calculation 3.docx")

            doc = DocxTemplate("LETTER.docx")
            context = {
                    'jan': monthly[12], 'jan_so': monthlyrad[12],
                    'feb': monthly[13], 'feb_so': monthlyrad[13],
                    'mar': monthly[14], 'mar_so': monthlyrad[14],
                    'apr': monthly[15], 'apr_so': monthlyrad[15],
                    'may': monthly[16], 'may_so': monthlyrad[16],
                    'jun': monthly[17], 'jun_so': monthlyrad[17],
                    'jul': monthly[18], 'jul_so': monthlyrad[18],
                    'aug': monthly[19], 'aug_so': monthlyrad[19],
                    'sep': monthly[20], 'sep_so': monthlyrad[20],
                    'oct': monthly[21], 'oct_so': monthlyrad[21],
                    'nov': monthly[22], 'nov_so': monthlyrad[22],
                    'dec': monthly[23], 'dec_so': monthlyrad[23],
                    'annual': annual[1], 'annual_so': annualrad[1],
                    'lat': location[0], 'lon': location[1], 'location': location_address,
                    'system_capacity': str(system_capacity_1) + ' kW',
                    'tilt': values[23][4], 'azimuth': values[24][4], 'losses': values[25][4],
                    'capacity_factor': capacity[1], 'range1': ranges_low[1], 'range2': ranges_high[1]

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

    data1 = (response_5.json())
    data2 = (response_6.json())

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
            jsonmonthly1 = data1['outputs']['ac_monthly']
            jsonmonthly2 = data2['outputs']['ac_monthly']
            jsonmonthlyrad1 = data1['outputs']['solrad_monthly']
            jsonmonthlyrad2 = data2['outputs']['solrad_monthly']
            jsonannual1 = data1['outputs']['ac_annual']
            jsonannual2 = data2['outputs']['ac_annual']
            jsonannualrad1 = data1['outputs']['solrad_annual']
            jsonannualrad2 = data2['outputs']['solrad_annual']
            lat = data1['station_info']['lat']
            lon = data2['station_info']['lon']
            jsoncap1 = data1['outputs']['capacity_factor']
            jsoncap2 = data2['outputs']['capacity_factor']
            
            rawannual = [jsonannual1, jsonannual2]
            rawannualrad = [jsonannualrad1, jsonannualrad2] 
            rawcap = [jsoncap1, jsoncap2]
            rawlocation = [lat, lon]

            monthly = []
            for x in jsonmonthly1:
                x = '{:0,.0f}'.format(x)
                monthly.append(x)
            for x in jsonmonthly2:
                x = '{:0,.0f}'.format(x)
                monthly.append(x)

            capacity = []
            for x in rawcap:
                x = '{:0,.1f}'.format(x)
                capacity.append(x)

            annualrad = []
            monthlyrad = []
            location = []
            for x in jsonmonthlyrad1:
                x = '{:0,.2f}'.format(x)
                monthlyrad.append(x)
            for x in jsonmonthlyrad2:
                x = '{:0,.2f}'.format(x)
                monthlyrad.append(x)
            for x in rawannualrad:
                x = '{:0,.2f}'.format(x)
                annualrad.append(x)
            for x in rawlocation:
                x = '{:0,.2f}'.format(x)
                location.append(x)

            solar_totals = [rawannual[0], rawannual[1]]
            low = [solar_totals[0]*0.024, solar_totals[1]*0.024]
            high = [solar_totals[0]*0.02218, solar_totals[1]*0.02218]
            low = [solar_totals[0] - low[0], solar_totals[1] - low[1]]
            high = [solar_totals[0] + high[0], solar_totals[1] + high[1]]

            ranges_low = []
            ranges_high = []
            annual = []
            for i in low:
                i = '{:0,.0f}'.format(i)
                ranges_low.append(i)
            for i in high:
                i = '{:0,.0f}'.format(i)
                ranges_high.append(i)
            for x in rawannual:
                x = '{:0,.0f}'.format(x)
                annual.append(x)
                    
            doc = DocxTemplate("LETTER.docx")
            context = {
                    'jan': monthly[0], 'jan_so': monthlyrad[0],
                    'feb': monthly[1], 'feb_so': monthlyrad[1],
                    'mar': monthly[2], 'mar_so': monthlyrad[2],
                    'apr': monthly[3], 'apr_so': monthlyrad[3],
                    'may': monthly[4], 'may_so': monthlyrad[4],
                    'jun': monthly[5], 'jun_so': monthlyrad[5],
                    'jul': monthly[6], 'jul_so': monthlyrad[6],
                    'aug': monthly[7], 'aug_so': monthlyrad[7],
                    'sep': monthly[8], 'sep_so': monthlyrad[8],
                    'oct': monthly[9], 'oct_so': monthlyrad[9],
                    'nov': monthly[10], 'nov_so': monthlyrad[10],
                    'dec': monthly[11], 'dec_so': monthlyrad[11],
                    'annual': annual[0], 'annual_so': annualrad[0],
                    'lat': location[0], 'lon': location[1], 'location': location_address,
                    'system_capacity': str(system_capacity_1) + ' kW',
                    'tilt': values[30][3], 'azimuth': values[31][3], 'losses': values[32][3],
                    'capacity_factor': capacity[0], 'range1': ranges_low[0], 'range2': ranges_high[0]

                }

            doc.render(context)
            if customer.pdf_response == 'YES':
                doc.save(customer.name + " PVWatts Calculation 5.docx")

                convert(f"{customer.name} PVWatts Calculation 5.docx", f"{customer.name} PVWatts Calculation 5.pdf")
                os.remove(f"{customer.name} PVWatts Calculation 5.docx")
            else:
                doc.save(customer.name + " PVWatts Calculation 5.docx")

            doc = DocxTemplate("LETTER.docx")
            context = {
                    'jan': monthly[12], 'jan_so': monthlyrad[12],
                    'feb': monthly[13], 'feb_so': monthlyrad[13],
                    'mar': monthly[14], 'mar_so': monthlyrad[14],
                    'apr': monthly[15], 'apr_so': monthlyrad[15],
                    'may': monthly[16], 'may_so': monthlyrad[16],
                    'jun': monthly[17], 'jun_so': monthlyrad[17],
                    'jul': monthly[18], 'jul_so': monthlyrad[18],
                    'aug': monthly[19], 'aug_so': monthlyrad[19],
                    'sep': monthly[20], 'sep_so': monthlyrad[20],
                    'oct': monthly[21], 'oct_so': monthlyrad[21],
                    'nov': monthly[22], 'nov_so': monthlyrad[22],
                    'dec': monthly[23], 'dec_so': monthlyrad[23],
                    'annual': annual[1], 'annual_so': annualrad[1],
                    'lat': location[0], 'lon': location[1], 'location': location_address,
                    'system_capacity': str(system_capacity_1) + ' kW',
                    'tilt': values[30][4], 'azimuth': values[31][4], 'losses': values[32][4],
                    'capacity_factor': capacity[1], 'range1': ranges_low[1], 'range2': ranges_high[1]

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
        letterOne()
    elif customer.array_count == 2:
        letterOne(), letterThree()
    elif customer.array_count == 3:
        letterOne(), letterTwo(), letterThree()
    else:
        print(f'ARRAY COUNT must be between 1-3')
        exit()

if __name__ == '__main__':
    start_time = time.time()
    main()