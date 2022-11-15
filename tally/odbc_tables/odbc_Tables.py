from tally.tally_connectivity import *

# List 'table_names' contains all the 191 table present in Tally
table_names = ['Ledger']


def odbc_columns(tablename):
    """
    Function to fetch all existing column names present in a table
    :param tablename: name of the table present in tally
    :return: list of all unique column names
    """
    cursor = tally_odbc()

    # Fetching column names by method cursor.columns(table)
    columns1 = [row.column_name.upper() for row in cursor.columns(table=tablename)]

    # Fetching column names by method data.description()
    query = f"select * from {tablename}"
    data = cursor.execute(query)
    columns2 = [(column[0]).upper() for column in data.description]

    return list(set(columns1 + columns2))


def odbc_table_data(location, tablename):
    """
    Converts the tally table data into csv file
    :param location: Path of the directory
    :param tablename: name of the table
    :return: Returns data of a particular table as a dataframe
    """
    table = []
    cursor = tally_odbc()
    columns = odbc_columns(tablename)

    # Executing sql query
    select = "select " + ','.join(j for j in columns) + " from " + tablename
    table_data = cursor.execute(select)

    col_list = []
    column = [column[0] for column in table_data.description]

    # Removing '$' or '_' from column names
    for cols in column:
        col_list.append(''.join(e for e in cols if e.isalnum()))
    col = tuple(col_list)

    # Appending retrieved data as a list of tuples
    for i in table_data:
        table.append(tuple(i))

    # Saving all the excel files at the given location
    file = tablename + ".csv"
    filename = os.path.join(location, file)
    print(filename)
    df = pd.DataFrame(table, columns=col)
    df.to_csv(filename, index=False)
    return df


def dataframe_to_csv(location):
    """
    Dumps data to a csv file of every table present in list 'table_names'
    :param location: Path of the directory
    :return: None
    """
    n = 1
    for i in table_names:
        odbc_table_data(location=location, tablename=i)
        print(n)
        n += 1



