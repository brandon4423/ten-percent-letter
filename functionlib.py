import json
import main
import creds
import requests
from docxtpl import DocxTemplate

def queryOne():
    query_1 = (f"&api_key={creds.api_key}&address={main.customer.address}&tilt={main.array_1.tilt}&azimuth={main.array_1.azimuth}&losses={main.array_1.losses}&module_type={main.customer.module_type}&array_type={main.customer.array_type}&system_capacity={main.system_capacity_1}")
    query_2 = (f"&api_key={creds.api_key}&address={main.customer.address}&tilt={main.array_2.tilt}&azimuth={main.array_2.azimuth}&losses={main.array_2.losses}&module_type={main.customer.module_type}&array_type={main.customer.array_type}&system_capacity={main.system_capacity_2}")
    response_1 = requests.get("https://developer.nrel.gov/api/pvwatts/v6.json?" + query_1)
    response_2 = requests.get("https://developer.nrel.gov/api/pvwatts/v6.json?" + query_2)

    data1 = (response_1.json())
    data2 = (response_2.json())

    with open('pvcal1.json', 'w') as file:
        json.dump(data1, file, indent=4)
    with open('pvcal2.json', 'w') as file:
        json.dump(data2, file, indent=4)

def queryTwo():
    query_3 = (f"&api_key={creds.api_key}&address={main.customer.address}&tilt={main.array_3.tilt}&azimuth={main.array_3.azimuth}&losses={main.array_3.losses}&module_type={main.customer.module_type}&array_type={main.customer.array_type}&system_capacity={main.system_capacity_3}")
    query_4 = (f"&api_key={creds.api_key}&address={main.customer.address}&tilt={main.array_4.tilt}&azimuth={main.array_4.azimuth}&losses={main.array_4.losses}&module_type={main.customer.module_type}&array_type={main.customer.array_type}&system_capacity={main.system_capacity_4}")
    response_3 = requests.get("https://developer.nrel.gov/api/pvwatts/v6.json?" + query_3)
    response_4 = requests.get("https://developer.nrel.gov/api/pvwatts/v6.json?" + query_4)

    data1 = (response_3.json())
    data2 = (response_4.json())

    with open('pvcal3.json', 'w') as file:
        json.dump(data1, file, indent=4)
    with open('pvcal4.json', 'w') as file:
        json.dump(data2, file, indent=4)

def queryThree():
    query_5 = (f"&api_key={creds.api_key}&address={main.customer.address}&tilt={main.array_5.tilt}&azimuth={main.array_5.azimuth}&losses={main.array_5.losses}&module_type={main.customer.module_type}&array_type={main.customer.array_type}&system_capacity={main.system_capacity_5}")
    query_6 = (f"&api_key={creds.api_key}&address={main.customer.address}&tilt={main.array_6.tilt}&azimuth={main.array_6.azimuth}&losses={main.array_6.losses}&module_type={main.customer.module_type}&array_type={main.customer.array_type}&system_capacity={main.system_capacity_6}")
    response_5 = requests.get("https://developer.nrel.gov/api/pvwatts/v6.json?" + query_5)
    response_6 = requests.get("https://developer.nrel.gov/api/pvwatts/v6.json?" + query_6)

    data1 = (response_5.json())
    data2 = (response_6.json())

    with open('pvcal5.json', 'w') as file:
        json.dump(data1, file, indent=4)
    with open('pvcal6.json', 'w') as file:
        json.dump(data2, file, indent=4)

def tenpercentOne():
    
    queryOne()
    with open('pvcal1.json', 'r') as file:
        data1 = json.load(file)

    with open('pvcal2.json', 'r') as file:
        data2 = json.load(file)

    json_data_1 = data1.copy()
    json_data_2 = data2.copy()
    json_data_1 = int(data1['outputs']['ac_annual'])
    json_data_2 = int(data2['outputs']['ac_annual'])
    difference = json_data_1 - json_data_2
    total = difference / json_data_1

    doc = DocxTemplate("TEN_PERCENT_V5.docx")
    context = {'hoa_name': main.customer.hoa_name, 'date': main.date, 'name': main.customer.name,
               'quantity': main.array_1.quantity, 'old_direction': main.array_1.direction, 'quantity2': main.array_2.quantity, 'state': main.customer.state,
               'old_azimuth': main.array_1.azimuth.replace("azimuth=", ""), 'old_tilt': main.array_1.tilt.replace("tilt=", ""), 'new_direction': main.array_2.direction,
               'new_azimuth': main.array_2.azimuth.replace("azimuth=", ""), 'new_tilt': main.array_2.tilt.replace("tilt=", ""), 
               'mod_watt': str(main.customer.mod_watt).replace("0.", "").strip(), 'percent': str(total)[2:][:2] + "%",
               'ac_monthly_original': str(json_data_1).split(main.ch, 1)[0], 'ac_monthly_new': str(json_data_2).split(main.ch, 1)[0]}

    doc.render(context)
    doc.save(main.customer.name + " Ten Percent Letter 1.docx")

def tenpercentTwo():
    queryTwo()
    with open('pvcal3.json', 'r') as file:
        data1 = json.load(file)

    with open('pvcal4.json', 'r') as file:
        data2 = json.load(file)

    json_data_3 = data1.copy()
    json_data_4 = data2.copy()
    json_data_3 = int(data1['outputs']['ac_annual'])
    json_data_4 = int(data2['outputs']['ac_annual'])
    difference = json_data_3 - json_data_4
    total = difference / json_data_3

    doc = DocxTemplate("TEN_PERCENT_V5.docx")
    context = {'hoa_name': main.customer.hoa_name, 'date': main.date, 'name': main.customer.name,
               'quantity': main.array_3.quantity, 'old_direction': main.array_3.direction, 'quantity2': main.array_4.quantity, 'state': main.customer.state,
               'old_azimuth': main.array_3.azimuth.replace("azimuth=", ""), 'old_tilt': main.array_3.tilt.replace("tilt=", ""), 'new_direction': main.array_4.direction,
               'new_azimuth': main.array_4.azimuth.replace("azimuth=", ""), 'new_tilt': main.array_4.tilt.replace("tilt=", ""), 
               'mod_watt': str(main.customer.mod_watt).replace("0.", "").strip(), 'percent': str(total)[2:][:2] + "%",
               'ac_monthly_original': str(json_data_3).split(main.ch, 1)[0], 'ac_monthly_new': str(json_data_4).split(main.ch, 1)[0]}

    doc.render(context)
    doc.save(main.customer.name + " Ten Percent Letter 2.docx")

def tenpercentThree():
    queryThree()
    with open('pvcal5.json', 'r') as file:
        data1 = json.load(file)

    with open('pvcal6.json', 'r') as file:
        data2 = json.load(file)

    json_data_5 = data1.copy()
    json_data_6 = data2.copy()
    json_data_5 = int(json_data_5['outputs']['ac_annual'])
    json_data_6 = int(json_data_6['outputs']['ac_annual'])
    difference = json_data_5 - json_data_6
    total = difference / json_data_5

    doc = DocxTemplate("TEN_PERCENT_V5.docx")
    context = {'hoa_name': main.customer.hoa_name, 'date': main.date, 'name': main.customer.name,
               'quantity': main.array_5.quantity, 'old_direction': main.array_5.direction, 'quantity2': main.array_6.quantity, 'state': main.customer.state,
               'old_azimuth': main.array_5.azimuth.replace("azimuth=", ""), 'old_tilt': main.array_5.tilt.replace("tilt=", ""), 'new_direction': main.array_6.direction,
               'new_azimuth': main.array_6.azimuth.replace("azimuth=", ""), 'new_tilt': main.array_6.tilt.replace("tilt=", ""), 
               'mod_watt': str(main.customer.mod_watt).replace("0.", "").strip(), 'percent': str(total)[2:][:2] + "%",
               'ac_monthly_original': str(json_data_5).split(main.ch, 1)[0], 'ac_monthly_new': str(json_data_6).split(main.ch, 1)[0]}

    doc.render(context)
    doc.save(main.customer.name + " Ten Percent Letter 3.docx")

def pvletterOne():
    with open('pvcal1.json', 'r') as file:
        data1 = json.load(file)

    with open('pvcal2.json', 'r') as file:
        data2 = json.load(file)

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
                'lat': location[0], 'lon': location[1], 'location': main.location_address,
                'system_capacity': str(main.system_capacity_1) + ' kW',
                'tilt': main.values[16][3], 'azimuth': main.values[17][3], 'losses': main.values[18][3],
                'capacity_factor': capacity[0], 'range1': ranges_low[0], 'range2': ranges_high[0]

            }

        doc.render(context)
        doc.save(main.customer.name + " PVWatts Calculation 1.docx")

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
                'lat': location[0], 'lon': location[1], 'location': main.location_address,
                'system_capacity': str(main.system_capacity_2) + ' kW',
                'tilt': main.values[16][4], 'azimuth': main.values[17][4], 'losses': main.values[18][4],
                'capacity_factor': capacity[1], 'range1': ranges_low[1], 'range2': ranges_high[1]

            }

        doc.render(context)
        doc.save(main.customer.name + " PVWatts Calculation 2.docx")

def pvletterTwo():
    with open('pvcal3.json', 'r') as file:
        data1 = json.load(file)

    with open('pvcal4.json', 'r') as file:
        data2 = json.load(file)

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
                'lat': location[0], 'lon': location[1], 'location': main.location_address,
                'system_capacity': str(main.system_capacity_3) + ' kW',
                'tilt': main.values[23][3], 'azimuth': main.values[24][3], 'losses': main.values[25][3],
                'capacity_factor': capacity[0], 'range1': ranges_low[0], 'range2': ranges_high[0]

            }

        doc.render(context)
        doc.save(main.customer.name + " PVWatts Calculation 3.docx")

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
                'lat': location[0], 'lon': location[1], 'location': main.location_address,
                'system_capacity': str(main.system_capacity_4) + ' kW',
                'tilt': main.values[23][4], 'azimuth': main.values[24][4], 'losses': main.values[25][4],
                'capacity_factor': capacity[1], 'range1': ranges_low[1], 'range2': ranges_high[1]

            }

        doc.render(context)
        doc.save(main.customer.name + " PVWatts Calculation 4.docx")

def pvletterThree():
    with open('pvcal5.json', 'r') as file:
        data1 = json.load(file)

    with open('pvcal6.json', 'r') as file:
        data2 = json.load(file)

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
                'lat': location[0], 'lon': location[1], 'location': main.location_address,
                'system_capacity': str(main.system_capacity_5) + ' kW',
                'tilt': main.values[30][3], 'azimuth': main.values[31][3], 'losses': main.values[32][3],
                'capacity_factor': capacity[0], 'range1': ranges_low[0], 'range2': ranges_high[0]

            }

        doc.render(context)
        doc.save(main.customer.name + " PVWatts Calculation 5.docx")

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
                    'lat': location[0], 'lon': location[1], 'location': main.location_address,
                    'system_capacity': str(main.system_capacity_6) + ' kW',
                    'tilt': main.values[30][4], 'azimuth': main.values[31][4], 'losses': main.values[32][4],
                    'capacity_factor': capacity[1], 'range1': ranges_low[1], 'range2': ranges_high[1]

            }

        doc.render(context)
        doc.save(main.customer.name + " PVWatts Calculation 6.docx")