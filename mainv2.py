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
location = values[12][0].replace(",,", ",")
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
class Pvdata:
        jan: float
        feb: float
        mar: float
        apr: float
        may: float
        jun: float
        jul: float
        aug: float
        sep: float
        oct: float
        nov: float
        dec: float
        annual: float
        lat: float
        lon: float
        jan_so: float
        feb_so: float
        mar_so: float
        apr_so: float
        may_so: float
        jun_so: float
        jul_so: float
        aug_so: float
        sep_so: float
        oct_so: float
        nov_so: float
        dec_so: float
        annual_so: float
        capacity_factor: float

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

if customer.state == 'TX':
    customer.state = tenref.texas
elif customer.state == 'CO':
    customer.state = tenref.colorado
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

    all_data_1 = (response_1.json())
    all_data_2 = (response_2.json())

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
            monthly_data_1 = all_data_1['outputs']['ac_monthly']
            monthly_data_2 = all_data_2['outputs']['ac_monthly']
            annual_1 = all_data_1['outputs']['ac_annual']
            annual_2 = all_data_2['outputs']['ac_annual']
            monthly_sodata_1 = all_data_1['outputs']['solrad_monthly']
            monthly_sodata_2 = all_data_2['outputs']['solrad_monthly']
            annualso_1 = all_data_1['outputs']['solrad_annual']
            annualso_2 = all_data_2['outputs']['solrad_annual']
            lat = all_data_1['station_info']['lat']
            lon = all_data_2['station_info']['lon']
            capacity_total_1 = all_data_1['outputs']['capacity_factor']
            capacity_total_2 = all_data_2['outputs']['capacity_factor']
            
            monthly_1 = Pvdata(monthly_data_1[0], monthly_data_1[1], monthly_data_1[2], monthly_data_1[3], monthly_data_1[4], monthly_data_1[5], monthly_data_1[6], monthly_data_1[7], monthly_data_1[8]
            , monthly_data_1[9], monthly_data_1[10], monthly_data_1[11], annual_1, lat, lon, monthly_sodata_1[0], monthly_sodata_1[1], monthly_sodata_1[2], monthly_sodata_1[3], monthly_sodata_1[4]
            , monthly_sodata_1[5], monthly_sodata_1[6], monthly_sodata_1[7], monthly_sodata_1[8], monthly_sodata_1[9], monthly_sodata_1[10], monthly_sodata_1[11], annualso_1, capacity_total_1)

            monthly_2 = Pvdata(monthly_data_2[0], monthly_data_2[1], monthly_data_2[2], monthly_data_2[3], monthly_data_2[4], monthly_data_2[5], monthly_data_2[6], monthly_data_2[7], monthly_data_2[8]
            , monthly_data_2[9], monthly_data_2[10], monthly_data_2[11], annual_2, lat, lon, monthly_sodata_2[0], monthly_sodata_2[1], monthly_sodata_2[2], monthly_sodata_2[3], monthly_sodata_2[4]
            , monthly_sodata_2[5], monthly_sodata_2[6], monthly_sodata_2[7], monthly_sodata_2[8], monthly_sodata_2[9], monthly_sodata_2[10], monthly_sodata_2[11], annualso_2, capacity_total_2)

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
            low = [solar_totals[0]*0.024, solar_totals[1]*0.024]
            high = [solar_totals[0]*0.02218, solar_totals[1]*0.02218]
            low = [solar_totals[0] - low[0], solar_totals[1] - low[1]]
            high = [solar_totals[0] + high[0], solar_totals[1] + high[1]]

            ranges_low = []
            for i in low:
                i = '{:0,.0f}'.format(i)
                ranges_low.append(i)

            ranges_high = []
            for i in high:
                i = '{:0,.0f}'.format(i)
                ranges_high.append(i)

            doc = DocxTemplate("LETTER.docx")
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

            doc = DocxTemplate("LETTER.docx")
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
                    'tilt': values[16][4], 'azimuth': values[17][4], 'losses': values[18][4],
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
            monthly_data_1 = all_data_1['outputs']['ac_monthly']
            monthly_data_2 = all_data_2['outputs']['ac_monthly']
            annual_1 = all_data_1['outputs']['ac_annual']
            annual_2 = all_data_2['outputs']['ac_annual']         
            monthly_sodata_1 = all_data_1['outputs']['solrad_monthly']
            monthly_sodata_2 = all_data_2['outputs']['solrad_monthly']
            annualso_1 = all_data_1['outputs']['solrad_annual']
            annualso_2 = all_data_2['outputs']['solrad_annual']
            lat = all_data_1['station_info']['lat']
            lon = all_data_2['station_info']['lon']
            capacity_total_1 = all_data_1['outputs']['capacity_factor']
            capacity_total_2 = all_data_2['outputs']['capacity_factor']
            
            monthly_1 = Pvdata(monthly_data_1[0], monthly_data_1[1], monthly_data_1[2], monthly_data_1[3], monthly_data_1[4], monthly_data_1[5], monthly_data_1[6], monthly_data_1[7], monthly_data_1[8]
            , monthly_data_1[9], monthly_data_1[10], monthly_data_1[11], annual_1, lat, lon, monthly_sodata_1[0], monthly_sodata_1[1], monthly_sodata_1[2], monthly_sodata_1[3], monthly_sodata_1[4]
            , monthly_sodata_1[5], monthly_sodata_1[6], monthly_sodata_1[7], monthly_sodata_1[8], monthly_sodata_1[9], monthly_sodata_1[10], monthly_sodata_1[11], annualso_1, capacity_total_1)

            monthly_2 = Pvdata(monthly_data_2[0], monthly_data_2[1], monthly_data_2[2], monthly_data_2[3], monthly_data_2[4], monthly_data_2[5], monthly_data_2[6], monthly_data_2[7], monthly_data_2[8]
            , monthly_data_2[9], monthly_data_2[10], monthly_data_2[11], annual_2, lat, lon, monthly_sodata_2[0], monthly_sodata_2[1], monthly_sodata_2[2], monthly_sodata_2[3], monthly_sodata_2[4]
            , monthly_sodata_2[5], monthly_sodata_2[6], monthly_sodata_2[7], monthly_sodata_2[8], monthly_sodata_2[9], monthly_sodata_2[10], monthly_sodata_2[11], annualso_2, capacity_total_2)

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
            low = [solar_totals[0]*0.024, solar_totals[1]*0.024]
            high = [solar_totals[0]*0.02218, solar_totals[1]*0.02218]
            low = [solar_totals[0] - low[0], solar_totals[1] - low[1]]
            high = [solar_totals[0] + high[0], solar_totals[1] + high[1]]

            ranges_low = []
            for i in low:
                i = '{:0,.0f}'.format(i)
                ranges_low.append(i)

            ranges_high = []
            for i in high:
                i = '{:0,.0f}'.format(i)
                ranges_high.append(i)

            doc = DocxTemplate("LETTER.docx")
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

            doc = DocxTemplate("LETTER.docx")
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
            monthly_data_1 = all_data_1['outputs']['ac_monthly']
            monthly_data_2 = all_data_2['outputs']['ac_monthly']
            annual_1 = all_data_1['outputs']['ac_annual']
            annual_2 = all_data_2['outputs']['ac_annual']           
            monthly_sodata_1 = all_data_1['outputs']['solrad_monthly']
            monthly_sodata_2 = all_data_2['outputs']['solrad_monthly']
            annualso_1 = all_data_1['outputs']['solrad_annual']
            annualso_2 = all_data_2['outputs']['solrad_annual']
            lat = all_data_1['station_info']['lat']
            lon = all_data_2['station_info']['lon']
            capacity_total_1 = all_data_1['outputs']['capacity_factor']
            capacity_total_2 = all_data_2['outputs']['capacity_factor']
            
            monthly_1 = Pvdata(monthly_data_1[0], monthly_data_1[1], monthly_data_1[2], monthly_data_1[3], monthly_data_1[4], monthly_data_1[5], monthly_data_1[6], monthly_data_1[7], monthly_data_1[8]
            , monthly_data_1[9], monthly_data_1[10], monthly_data_1[11], annual_1, lat, lon, monthly_sodata_1[0], monthly_sodata_1[1], monthly_sodata_1[2], monthly_sodata_1[3], monthly_sodata_1[4]
            , monthly_sodata_1[5], monthly_sodata_1[6], monthly_sodata_1[7], monthly_sodata_1[8], monthly_sodata_1[9], monthly_sodata_1[10], monthly_sodata_1[11], annualso_1, capacity_total_1)

            monthly_2 = Pvdata(monthly_data_2[0], monthly_data_2[1], monthly_data_2[2], monthly_data_2[3], monthly_data_2[4], monthly_data_2[5], monthly_data_2[6], monthly_data_2[7], monthly_data_2[8]
            , monthly_data_2[9], monthly_data_2[10], monthly_data_2[11], annual_2, lat, lon, monthly_sodata_2[0], monthly_sodata_2[1], monthly_sodata_2[2], monthly_sodata_2[3], monthly_sodata_2[4]
            , monthly_sodata_2[5], monthly_sodata_2[6], monthly_sodata_2[7], monthly_sodata_2[8], monthly_sodata_2[9], monthly_sodata_2[10], monthly_sodata_2[11], annualso_2, capacity_total_2)

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
            low = [solar_totals[0]*0.024, solar_totals[1]*0.024]
            high = [solar_totals[0]*0.02218, solar_totals[1]*0.02218]
            low = [solar_totals[0] - low[0], solar_totals[1] - low[1]]
            high = [solar_totals[0] + high[0], solar_totals[1] + high[1]]

            ranges_low = []
            for i in low:
                i = '{:0,.0f}'.format(i)
                ranges_low.append(i)

            ranges_high = []
            for i in high:
                i = '{:0,.0f}'.format(i)
                ranges_high.append(i)

            doc = DocxTemplate("LETTER.docx")
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

            doc = DocxTemplate("LETTER.docx")
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