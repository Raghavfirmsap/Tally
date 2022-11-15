print("__________")

import os
from tally.odbc_tables.odbc_Tables import dataframe_to_csv
from tally.tally_connectivity import tally_odbc
from tally.xml_reports.Reports import all_reports
import time
import sys

print("____________")


def main():
    """
    Creates a directory at specified path and specified name.
    Creates 197 csv files.
    :return: None
    """
    start = time.time()
    path = sys.argv[1]
    startdate = sys.argv[2]
    enddate = sys.argv[3]
    filename = sys.argv[4]
    print(path, startdate, enddate, filename)

    main_dir = path + r"\TallyData_" + filename
    os.mkdir(main_dir, mode=0o666)

    odbc = main_dir + r"\odbc"
    os.mkdir(odbc, mode=0o666)
    dataframe_to_csv(location=odbc)

    xml = main_dir + r"\xml"
    os.mkdir(xml, mode=0o666)
    all_reports(location=xml, startdate=startdate, enddate=enddate)

    end = time.time()
    print("Directory '% s' is built with tally data!!!!" % main_dir)
    print("Time taken:", end - start)


if __name__ == '__main__':
    print("Reached")
    print("name: " + sys.argv[1])
    print("name: " + sys.argv[2])
    print("name: " + sys.argv[3])
    print("name: " + sys.argv[4])
    main()
