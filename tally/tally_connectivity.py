import csv
import datetime
import json
import os
from pprint import pprint

import pyodbc
import requests
import xmltodict


def tally_request(parameter):
    """
    Establishing connectivity with tally by xml method
    :param parameter: XML format
    :return: response data
    """
    url = "http://127.0.0.1:9000"
    headers = {"Content-type": "text/xml;charset=UTF-8", "Accept": "text/xml"}
    response = requests.post(url=url, data=parameter, headers=headers, stream=True)
    print("sending response")
    # print("_______________", response.content)
    # print(response.content)
    return response


def tally_odbc():
    """
    Establishing connectivity with tally by odbc method
    :return: Returns cursor object
    """
    odbc_list = ["ODBC", "ODBC32", "ODBC64"]
    for odbc in odbc_list:
        try:
            odbc = 'DSN=Tally%s_%s' % (odbc, 9000)
            conn = pyodbc.connect(odbc)
            cursor = conn.cursor()
            break
        except pyodbc.InterfaceError as ex:
            pass
    else:
        raise ValueError("Connection with tally is not successful")
    return cursor


def flatten_json(y):
    """
    Flattens a deeply nested dictionary
    :param y: Nested dictionary containing lists and dictionaries
    :return: Returns a flatten dictionary
    """
    out = {}

    def flatten(x, name=''):
        if type(x) is dict:
            for a in x:
                flatten(x[a], name + a + '_')
        elif type(x) is list:
            i = 0
            for a in x:
                flatten(a, name + str(i) + '_')
                i += 1
        else:
            out[name[:-1]] = x

    flatten(y)

    return out
