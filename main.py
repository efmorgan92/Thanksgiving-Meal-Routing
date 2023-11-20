from openpyxl import Workbook
import openpyxl
import pandas as pd
from math import radians, sin, cos, sqrt, atan2
from geopy.distance import geodesic
from ast import literal_eval
import googlemaps
import xlsxwriter
from openpyxl.utils.dataframe import dataframe_to_rows
import os

API_KEY = 'YOUR API KEY HERE'
SHEET_NAME = '../2023 TG Routing Input v2.xlsx'

sheet_names = pd.ExcelFile(SHEET_NAME).sheet_names

location = []
driver_num = []
for sheet in sheet_names:
    location.append(sheet.split(' - ')[0])
    driver_num.append(int(sheet.split(' - ')[1].split(' ')[0]))

#
# def get_lat_lon(address):
#
#     # Geocode the address
#     geocode_result = gmaps.geocode(address)
#
#     if not geocode_result:
#         print(f"Could not find coordinates for the address: {address}")
#         return None
#
#     # Extract the latitude and longitude
#     location = geocode_result[0]['geometry']['location']
#     lat, lon = location['lat'], location['lng']
#
#     return [lat, lon]
#
#
# gmaps = googlemaps.Client(key=API_KEY)
# for sheet in range(len(location)):
#     df = pd.read_excel(SHEET_NAME, sheet_name=sheet)
#
#     df['Full Address'] = df.apply(lambda row: ', '.join([str(row['Street Address']), str(row['City']), str(row['Zip Code'])]), axis=1)
#     df['Lat/Lon'] = df.apply(lambda row: get_lat_lon(str(row['Full Address'])), axis=1)
#
#     output_file = 'Output_' + str(location[sheet]) + '.xlsx'
#     df.to_excel(output_file, index=False)


def haversine(lat1, lon1, lat2, lon2):
    # Using the geopy library for more accurate distance calculation
    coord1 = (lat1, lon1)
    coord2 = (lat2, lon2)
    return geodesic(coord1, coord2).kilometers

def sort_coordinates_by_proximity(coordinates):
    # Assume coordinates is a list of tuples (latitude, longitude)
    sorted_coordinates = sorted(coordinates, key=lambda coord: coord[0])  # Sort by latitude

    # You can further refine the sorting logic based on proximity using the haversine distance
    sorted_coordinates = sorted(coordinates, key=lambda coord: haversine(coordinates[0][0], coordinates[0][1], coord[0], coord[1]))

    return sorted_coordinates


output_files = []
for sheet in range(len(location)):
    df = pd.read_excel('Output_' + str(location[sheet]) + '.xlsx')
    coordinates = df['Lat/Lon'].to_list()
    coord_tuples = []

    for coord in coordinates:
        coord = literal_eval(coord)
        coord_tuples.append((coord[0], coord[1]))

    # Sort coordinates by proximity
    sorted_coordinates = sort_coordinates_by_proximity(coord_tuples)

    df['Order'] = [sorted_coordinates.index(value) for value in coord_tuples]
    df_by_order = df.sort_values(by='Order')

    min_stops = df['# Meals'].count() // int(driver_num[sheet])
    extra_stops = df['# Meals'].count() % int(driver_num[sheet])

    route = []
    route_num = 1
    stop_count = 1
    use_stop = True

    for item in df['# Meals']:
        route.append(route_num)

        if stop_count == min_stops:
            if extra_stops > 0 and use_stop:
                extra_stops -= 1
                use_stop = False
            else:
                route_num += 1
                stop_count = 1
                use_stop = True
        else:
            stop_count += 1


    df_by_order['Route'] = route
    df_by_order = df_by_order.drop(['Full Address', 'Lat/Lon', 'Order'], axis=1)
    output_files.append(df_by_order)

    # Create the relevant folder

    folder_name = f'{location[sheet]}'
    os.makedirs(folder_name, exist_ok=True)

    # Create the individual route files
    for driver in range(driver_num[sheet]):

        current_route = driver + 1

        route_df = df_by_order[df_by_order['Route'] == current_route]
        route_df = route_df.drop(['Parent or Legal Guardian\'s Name (if applicable)', 'CLIENT CODE ','Route'], axis=1)

        # Read the existing Excel file
        input_file = 'route_detail_template.xlsx'
        new_workbook = openpyxl.load_workbook(input_file)

        # Select the desired sheet
        workbook = new_workbook['Sheet1']  # Replace 'Sheet1' with the actual sheet name

        workbook['A2'] = 'Route #:'
        workbook['A3'] = 'Location: '
        workbook['B2'] = f'{current_route}'
        workbook['B3'] = f'{location[sheet]}'
        workbook['I2'] = 'Total Meals: '
        workbook['I3'] = 'Total Stops'
        workbook['J2'] = f'{route_df["# Meals"].sum()}'
        workbook['J3'] = f'{route_df["# Meals"].count()}'
        workbook['A4'] = ' '

        start_cell = 'A6'

        data_values = route_df.values.tolist()
        for row_index, row in enumerate(data_values):
            for col_index, value in enumerate(row):
                workbook[start_cell].offset(row=row_index, column=col_index).value = value

        if location[sheet] == 'FUMC CG':
            route_df['Singles'] = route_df['# Meals'] % 2
            route_df['Doubles'] = route_df['# Meals'] // 2
            singles = route_df['Singles'].sum()
            doubles = route_df['Doubles'].sum()

            workbook['I15'] = f'{doubles} double(s)'
            workbook['I16'] = f'{singles} single(s)'

        # Save the Excel workbook
        file_name = f'./{location[sheet]}/{location[sheet]}_Route {current_route}.xlsx'
        new_workbook.save(file_name)

    # Create the master route list
    input_file = 'route_master_template.xlsx'
    new_workbook = openpyxl.load_workbook(input_file)

    # Select the desired sheet
    workbook = new_workbook['Sheet1']  # Replace 'Sheet1' with the actual sheet name
    workbook['B1'] = f'{location[sheet]}'

    # Group by 'Category' and calculate count and sum
    df_master = df_by_order.groupby('Route').agg({'# Meals': ['count','sum']}).reset_index()
    df_master.columns = ['Category', 'RowCount', 'SumOfQuantity']
    df_master['NewColumn1'] = " "
    df_master['NewColumn2'] = " "
    df_master = df_master[['Category', 'NewColumn1', 'NewColumn2', 'RowCount', 'SumOfQuantity']]

    start_cell = 'A3'
    data_values = df_master.values.tolist()
    for row_index, row in enumerate(data_values):
        for col_index, value in enumerate(row):
            workbook[start_cell].offset(row=row_index, column=col_index).value = value

    # Save the Excel workbook
    file_name = f'./{location[sheet]}/{location[sheet]}_Master Route List.xlsx'
    new_workbook.save(file_name)









