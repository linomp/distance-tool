import time
import urllib.parse
from io import BytesIO

import pandas as pd
import requests


class GoogleMapsRequestException(Exception):
    pass


def get_distance_in_km(source: str, dest: str, api_key: str):
    # Source: https://developers.google.com/maps/documentation/distance-matrix/distance-matrix#maps_http_distancematrix_latlng-py

    source = urllib.parse.quote(source)
    dest = urllib.parse.quote(dest)

    url = f"https://maps.googleapis.com/maps/api/distancematrix/json?mode=driving&origins={source}&destinations={dest}&units=metric&key={api_key}"

    payload = {}
    headers = {}

    response = requests.request("GET", url, headers=headers, data=payload)

    response = response.json()

    if "error_message" in response:
        raise GoogleMapsRequestException(response["error_message"])

    distance: float = response['rows'][0]['elements'][0]['distance']['value']

    # convert from m to km
    distance /= 1000

    return distance


def process_input_file(input_file: str | BytesIO, api_key: str, standalone_mode: bool = False):
    """
    Function that reads the 2 first columns of an excel file, and calls the get_distance_in_km function, then writes the result in a new column, and saves the file as a new excel file
    """

    df = pd.read_excel(input_file, engine='openpyxl', header=None)

    df = df.iloc[:, :2]

    for index, row in df.iterrows():
        time.sleep(0.005)
        try:
            distance = get_distance_in_km(row[0], row[1], api_key)
        except GoogleMapsRequestException as e:
            raise e
        except Exception:
            distance = 'error'

        # write distance in 3rd column
        df.loc[index, 'distance'] = distance

    # output_file_name is the same as input but "with_distances" and the current time in hh:mm:ss
    output_file_name = 'computed_distances_' + time.strftime("%H_%M_%S") + '.xlsx'

    if standalone_mode:
        df.to_excel(output_file_name, index=False, header=["Source", "Destination", "Distance (km)"])
    else:
        return to_excel(df, header=["Source", "Destination", "Distance (km)"]), output_file_name


def to_excel(df, header=None):
    output = BytesIO()
    writer = pd.ExcelWriter(output)
    df.to_excel(writer, index=False, header=header)
    writer.close()
    processed_data = output.getvalue()
    return processed_data


def load_api_key():
    with open('.env', 'r') as f:
        key = f.read().split('=')[1]
        if not key:
            raise Exception('No API key found')
    return key


if __name__ == '__main__':
    api_key = load_api_key()
    process_input_file("input.xlsx", api_key=api_key, standalone_mode=True)
