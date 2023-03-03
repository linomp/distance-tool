import os
import time
from io import BytesIO

import requests
import urllib.parse
import pandas as pd


def get_distance_in_km(source: str, dest: str, api_key: str):
    source = urllib.parse.quote(source)
    dest = urllib.parse.quote(dest)

    url = f"https://maps.googleapis.com/maps/api/distancematrix/json?origins={source}&destinations={dest}&units=metric&key={api_key}"

    payload = {}
    headers = {}

    response = requests.request("GET", url, headers=headers, data=payload)

    response = response.json()

    distance = response['rows'][0]['elements'][0]['distance']['text']

    # remove "km" from the string
    distance = distance[:-2]

    return distance


def load_api_key():
    key = os.getenv('GOOGLE_MAPS_API_KEY')
    if not key:
        with open('.env', 'r') as f:
            key = f.read().split('=')[1]

            if not key:
                raise Exception('No API key found')
    return key


def process_input_file(input_file: str | BytesIO, api_key: str | None = None):
    """
    Function that reads the 2 first columns of an excel file, and calls the get_distance_in_km function, then writes the result in a new column, and saves the file as a new excel file
    """

    if api_key is None:
        api_key = load_api_key()

    df = pd.read_excel(input_file, engine='openpyxl', header=None)

    df = df.iloc[:, :2]

    for index, row in df.iterrows():
        time.sleep(0.1)
        try:
            distance = get_distance_in_km(row[0], row[1], api_key)
        except:
            distance = 'error'

        # write distance in 3rd column
        df.loc[index, 'distance'] = distance

    # output_file_name is the same as input but "with_distances" and the current time in hh:mm:ss
    output_file_name = 'computed_distances_' + time.strftime("%H_%M_%S") + '.xlsx'

    # df.to_excel(output_file_name, index=False, header=["Source", "Destination", "Distance (km)"])

    return to_excel(df, header=["Source", "Destination", "Distance (km)"]), output_file_name


def to_excel(df, header=None):
    output = BytesIO()
    writer = pd.ExcelWriter(output)
    df.to_excel(writer, index=False, header=header)
    writer.close()
    processed_data = output.getvalue()
    return processed_data


if __name__ == '__main__':
    process_input_file("input.xlsx")
