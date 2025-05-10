import pandas as pd
from geopy.geocoders import Bing
import openpyxl  # needed to be able to stream df directly to execel files
from geopy.extra.rate_limiter import RateLimiter

# Load Excel file
file_path = "./data/Donation_Reg_21-22.xlsx"
sheet_name = "members"
address_column = "Address"

df = pd.read_excel(file_path, sheet_name=sheet_name)

api_key = "Am1g2QGAM515_5-QhOJe4-P-vcJGSBoSMIrYV3e_gBbzI50ihxjSYEC1_sPuHntr"
geolocator = Bing(api_key)

# So as to not spam the API server
geocode = RateLimiter(geolocator.geocode, min_delay_seconds=1)


def get_lat_lng(address):
    try:
        location = geocode(address)
        if location:
            return location.latitude, location.longitude
        else:
            return None, None
    except Exception as e:
        print(f"Error: {e}")
        return None, None


l = []
for i in range(23):
    l.append(get_lat_lng(df[address_column][i]))


df1 = pd.DataFrame(l, columns=["latitude", "longitude"])
df1.to_excel("output.xlsx")

# # Add 'Latitude' and 'Longitude' columns to the DataFrame
# df["Latitude"], df["Longitude"] = zip(*df[address_column].apply(get_lat_lng))

# # Save the updated DataFrame to a new Excel file
# output_file_path = "your_output_file.xlsx"
# df.to_excel(output_file_path, index=False)
