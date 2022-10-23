import time
import tenref
import gspread
from dataclasses import dataclass

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
system_capacity_3 = customer.mod_watt * array_3.quantity
system_capacity_4 = customer.mod_watt * array_4.quantity
system_capacity_5 = customer.mod_watt * array_5.quantity
system_capacity_6 = customer.mod_watt * array_6.quantity

if customer.state == 'TX':
    customer.state = tenref.texas
elif customer.state == 'CO':
    customer.state = tenref.colorado
else:
    pass

def main():
    start_time = time.time()
    from functionlib import tenpercentOne, tenpercentTwo, tenpercentThree
    from functionlib import pvletterOne, pvletterTwo, pvletterThree
    if customer.array_count == 1:
        tenpercentOne()
        if customer.pvwatts == 'YES':
            pvletterOne()
    elif customer.array_count == 2:
        tenpercentOne(), tenpercentTwo()
        if customer.pvwatts == 'YES':
            pvletterOne(), pvletterTwo()
    elif customer.array_count == 3:
        tenpercentOne(), tenpercentTwo(), tenpercentThree()
        if customer.pvwatts == 'YES':
            pvletterOne(), pvletterTwo(), pvletterThree()
    else:
        print(f'ARRAY COUNT must be between 1-3')
        exit()

    end_time = time.time()
    final_time = end_time - start_time
    print(f"Run in: {final_time} seconds")

if __name__ == '__main__':
    start_time = time.time()
    main()