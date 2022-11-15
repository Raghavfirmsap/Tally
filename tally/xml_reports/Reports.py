from tally.tally_connectivity import *
from tally.xml_reports.Voucher_Register import vouchers
global company_name
global dates
import xml.etree.ElementTree as ET




def balance_sheet(location):
    """
    programmatically fetch Balance sheet details
    :param location: Path of the directory
    :return: None
    """
    try:
        params = """
                    <ENVELOPE>
                    <HEADER>
                    <ID>Balance Sheet</ID>
                    <VERSION>1</VERSION>
                    <TALLYREQUEST>EXPORT</TALLYREQUEST>
                    <TYPE>DATA</TYPE>
                    </HEADER>
                    <BODY>
                    <DESC>
                    <STATICVARIABLES>
                    <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
                    </STATICVARIABLES>
                    </DESC>
                    </BODY>
                    </ENVELOPE>
                """

        response = tally_request(parameter=params)
        response = xmltodict.parse(response.content)
        response = json.loads(json.dumps(response['ENVELOPE']))

        name = pd.json_normalize(response['BSNAME'])
        amt = pd.json_normalize(response['BSAMT'])
        final = name.merge(amt, left_index=True, right_index=True)

        filename = os.path.join(location, "BalanceSheet.csv")
        final.to_csv(filename, index=False)
        print(filename)
    except Exception as e:
        print(e)


def list_account(location):
    """
    programmatically fetch List of Accounts details for Ledger and Groups
    :param location:
    :return: Path of the directory
    """
    try:
        group_data = []
        ledger_data = []
        params = """
                        <ENVELOPE>
                        <HEADER>
                        <VERSION>1</VERSION>
                        <TALLYREQUEST>EXPORT</TALLYREQUEST>
                        <TYPE>DATA</TYPE>
                        <ID>List of Accounts</ID>
                        </HEADER>
                        <BODY>
                        <DESC>
                        <STATICVARIABLES>
                        <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
                        </STATICVARIABLES>
                        </DESC>
                        </BODY>
                        </ENVELOPE>
                    """

        response = tally_request(parameter=params)
        response = response.content
        response = response.decode('utf-8', 'ignore')
        response = response.replace("&#4;", "")
        response = response.encode('utf-8')
        response = xmltodict.parse(response)
        response = json.loads(json.dumps(response))
        global company_name
        company_name = response['ENVELOPE']['BODY']['DESC']['STATICVARIABLES']['SVCURRENTCOMPANY']
        response = response['ENVELOPE']['BODY']['DATA']['TALLYMESSAGE']
        print(company_name)
        for i in range(len(response)):
            if 'LEDGER' in response[i].keys():
                res = response[i]['LEDGER']
                flat = flatten_json(res)
                ledger_data.append(flat)
            elif 'GROUP' in response[i].keys():
                res = response[i]['GROUP']
                flat = flatten_json(res)
                group_data.append(flat)
        ledger_df = pd.DataFrame(ledger_data)
        group_df = pd.DataFrame(group_data)
        ledger_filename = os.path.join(location, "Ledger_ListofAccounts.csv")
        group_filename = os.path.join(location, "Group_ListofAccounts.csv")
        ledger_df.to_csv(ledger_filename, index=False)
        print(ledger_filename)
        group_df.to_csv(group_filename, index=False)
        print(group_filename)
    except Exception as e:
        print(e)


def profit_loss(location):
    """
    programmatically fetch Profit and Loss A/c details
    :param location: Path of the directory
    :return: None
    """
    try:
        param = '''
                        <ENVELOPE>
                        <HEADER>
                        <TALLYREQUEST>Export Data</TALLYREQUEST>
                        </HEADER>
                        <BODY>
                        <EXPORTDATA>
                        <REQUESTDESC>
                        <STATICVARIABLES>
                        <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
                        <DSPSHOWOPENING>YES</DSPSHOWOPENING>
                        <DSPSHOWTRANS>YES</DSPSHOWTRANS>
                        <EXPLODEALLLEVELS>YES</EXPLODEALLLEVELS>
                        <EXPLODEFLAG>YES</EXPLODEFLAG>
                        </STATICVARIABLES>
                        <REPORTNAME>Profit and Loss</REPORTNAME>
                        </REQUESTDESC>
                        </EXPORTDATA>
                        </BODY>
                        </ENVELOPE>
                    '''
        response = tally_request(parameter=param)
        response = xmltodict.parse(response.content)
        response = json.loads(json.dumps(response['ENVELOPE']))

        for sub_dict in response['BSAMT']:
            sub_dict['SUBAMT'] = sub_dict.pop('BSSUBAMT')
            sub_dict['MAINAMT'] = sub_dict.pop('BSMAINAMT')
        for sub_dict in response['PLAMT']:
            sub_dict['SUBAMT'] = sub_dict.pop('PLSUBAMT')
            sub_dict['MAINAMT'] = sub_dict.pop('BSMAINAMT')

        name1 = pd.json_normalize(response['DSPACCNAME'])
        amt1 = pd.json_normalize(response['PLAMT'])
        final1 = name1.merge(amt1, left_index=True, right_index=True)
        name2 = pd.json_normalize(response['BSNAME'])
        amt2 = pd.json_normalize(response['BSAMT'])
        final2 = name2.merge(amt2, left_index=True, right_index=True)

        final1.rename(columns={'DSPDISPNAME': 'NAME'}, inplace=True)
        final2.rename(columns={'DSPACCNAME.DSPDISPNAME': 'NAME'}, inplace=True)

        mod_df = final1.append(final2, ignore_index=True)
        filename = os.path.join(location, "Profit&Loss.csv")
        mod_df.to_csv(filename, index=False)
        print(filename)
    except Exception as e:
        print(e)


def trial_balance_ledger_wise(location, startdate = dates[0], enddate = dates[1]):
    """
    programmatically fetch Trial Balance details
    :param location: Path of the directory
    :return: None
    """
    try:
        params = f"""
                        <ENVELOPE>
                        <HEADER>
                        <ID>Trial Balance</ID>
                        <VERSION>1</VERSION>
                        <TALLYREQUEST>EXPORT</TALLYREQUEST>
                        <TYPE>DATA</TYPE>
                        </HEADER>
                        <BODY>

                        <DESC>
                        <STATICVARIABLES>
                        <SVFROMDATE>{startdate}</SVFROMDATE>
                        <SVTODATE>{enddate}</SVTODATE>
                        <DSPSHOWOPENING>YES</DSPSHOWOPENING>
                        <DSPSHOWTRANS>YES</DSPSHOWTRANS>
                        <EXPLODEALLLEVELS>YES</EXPLODEALLLEVELS>
                        <EXPLODEFLAG>YES</EXPLODEFLAG>
                        <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
                        </STATICVARIABLES>
                        <TDL>
                            <TDLMESSAGE>
                                <REPORT NAME="TrialBalance" ISMODIFY="Yes">
                                    <ADD>  SET: IsLedgerWise :"Yes"</ADD>
                                </REPORT>
                            </TDLMESSAGE>
                        </TDL>
                        </DESC>
                        </BODY>
                        </ENVELOPE>
                    """

        response = tally_request(parameter=params)
        # response = xmltodict.parse(response.content)
        # response = ET.parse(response.content, parser=ET.XMLParser(encoding='iso-8859-5'))
        response = response.content
        response = response.decode('utf-8', 'ignore')
        response = response.replace("&#4;", "")
        response = response.encode('utf-8')
        response = xmltodict.parse(response)
        print("Reached")
        response = json.loads(json.dumps(response['ENVELOPE']))

        name = pd.json_normalize(response['DSPACCNAME'])
        amt = pd.json_normalize(response['DSPACCINFO'])
        final = name.merge(amt, left_index=True, right_index=True)
        filename = os.path.join(location, "TrialBalance.csv")
        final.to_csv(filename, index=False)
        print(filename)
    except Exception as e:
        print(e)


def all_reports(location, startdate, enddate):
    global company_name

    try:
        cursor = tally_odbc()
        print(company_name)
        query = f"select $Name, $StartingFrom, $AuditedUpto from Company"
        data = cursor.execute(query)
        global dates
        dates = [row for row in data if company_name in row]
        print(dates)
    except Exception as e:
        print(e)

    trial_balance_ledger_wise(location, startdate, enddate)
    profit_loss(location)
    balance_sheet(location)
    list_account(location)
    print("Rest is done")
    vouchers(location=location, CompanyName=company_name, startdate=startdate, enddate=enddate)
