from tally.tally_connectivity import *
from datetime import timedelta
import time
import datetime
import xml.etree.ElementTree as ET
from tally.xml_reports.Reports import *


keys = ["@VCHKEY", "@VCHTYPE", "DATE", "GUID", "STATENAME", "NARRATION", "COUNTRYOFRESIDENCE", "PARTYGSTIN",
        "NATUREOFSALES", "PARTYNAME", "PARTYLEDGERNAME", "VOUCHERTYPENAME", "REFERENCE", "VOUCHERNUMBER",
        "BASICBASEPARTYNAME", "CONSIGNEEGSTIN", "BASICBUYERNAME", "BUYERPINNUMBER", "CONSIGNEESTATENAME",
        "CONSIGNEEPINCODE", "EFFECTIVEDATE", "ISDELETED", "ISCANCELLED", "HASCASHFLOW", "ALTERID", "MASTERID"]

col = keys + ["LEDGERNAME", "AMOUNT", "Stock Item", "Quantity", "Rate"]
# col = keys + ["LEDGERNAME", "AMOUNT"]


def vouchers(location, CompanyName,startdate = dates[0], enddate = dates[1]):
    """
    programmatically fetch Voucher details
    :param location: Path of the directory
    :param CompanyNumber: unique company number (example -> 100002)
    :return: None
    """
    try:
        cursor = tally_odbc()
        print(CompanyName)
        query = f"select $Name, $StartingFrom, $AuditedUpto from Company"
        data = cursor.execute(query)
        dates = [row for row in data if CompanyName in row]
        print(dates)
        start = datetime.date(int(startdate[:4]),int(startdate[4:6]),int(startdate[6:]))
        end = datetime.date(int(enddate[:4]),int(enddate[4:6]),int(enddate[6:]))
        print(start)
        n = 1
        filename = os.path.join(location, "Voucher.csv")
        with open(filename, 'w', newline='') as f:
            writer = csv.DictWriter(f, fieldnames=col)
            writer.writeheader()
            while start <= end:
                print("___________")
                print(start)
                new_end = start + timedelta(4)
                print(new_end)
                params = """
                                <ENVELOPE>
                                <HEADER>
                                <VERSION>1</VERSION>
                                <TALLYREQUEST>EXPORT</TALLYREQUEST>
                                <TYPE>DATA</TYPE>
                                <ID>Voucher Register</ID>
                                </HEADER>
                                <BODY>
                                <DESC>
                                <STATICVARIABLES>
                                <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
                                <SVFROMDATE>{}</SVFROMDATE>
                                <SVTODATE>{}</SVTODATE>
                                </STATICVARIABLES>
                                </DESC>
                                </BODY>
                                </ENVELOPE>
                            """
                params = params.format(start.strftime('%d-%m-%Y'), new_end.strftime('%d-%m-%Y'))
                response = tally_request(parameter=params)
                response = response.content

                response = response.decode('utf-8', "ignore")
                response = response.replace("&#4;", "")
                response = response.replace("&#13;&#10;", "")
                response = response.replace("Ͽ?Ͼ", " ")
                response = response.replace("", " ")
                # print(response)
                response = response.encode('utf-8')

                try:
                    response = xmltodict.parse(response)
                except Exception as e:
                    print(response)
                    print(e)
                    raise

                response = json.loads(json.dumps(response['ENVELOPE']['BODY']['DATA']))
                with open("out.txt", "w", encoding="utf-8") as f:
                    f.write(str(response))

                if response is not None:
                    response = response['TALLYMESSAGE']
                    print("res_____________",response)

                    if type(response) == list:
                        # print("response is list")
                        for v in range(len(response)):
                            if "VOUCHER" in response[v]:
                                res = response[v]['VOUCHER']
                                # print("Voucher")
                                # print("___________")
                                if "ALLLEDGERENTRIES.LIST" in res and res["ALLLEDGERENTRIES.LIST"] is not None:
                                    a = res["ALLLEDGERENTRIES.LIST"]
                                    # print("ALLLEDGERENTRIES")
                                    # print(a)
                                    if type(a) == dict:
                                        # print("a___________dict")
                                        filtered_res = dict((k, res[k]) for k in keys)
                                        filtered_res['LEDGERNAME'] = a.get('LEDGERNAME', None)
                                        filtered_res['AMOUNT'] = a.get('AMOUNT', None)
                                        #
                                        # print(filtered_res['LEDGERNAME'])
                                        # print(filtered_res['AMOUNT'])
                                    else:
                                        for i in res['ALLLEDGERENTRIES.LIST']:
                                            # print("a___________list")
                                            # print(i)
                                            filtered_res = dict((k, res[k]) for k in keys)
                                            # print(filtered_res)
                                            filtered_res['LEDGERNAME'] = i.get('LEDGERNAME', None)
                                            filtered_res['AMOUNT'] = i.get('AMOUNT', None)
                                            # print(filtered_res['LEDGERNAME'])
                                            # print(filtered_res['AMOUNT'])
                                            writer.writerow(filtered_res)
                                            # print("_____________")

                                if "LEDGERENTRIES.LIST" in res and res["LEDGERENTRIES.LIST"] is not None:
                                    b = res["LEDGERENTRIES.LIST"]

                                    # print("b_______________",b)
                                    if type(b) == list:
                                        for i in res['LEDGERENTRIES.LIST']:
                                            filtered_res = dict((k, res[k]) for k in keys)
                                            filtered_res['LEDGERNAME'] = i.get('LEDGERNAME', None)
                                            filtered_res['AMOUNT'] = i.get('AMOUNT', None)
                                            # print(filtered_res['LEDGERNAME'])
                                            # print(filtered_res['AMOUNT'])
                                            writer.writerow(filtered_res)
                                    elif type(b) == dict:
                                        filtered_res = dict((k, res[k]) for k in keys)
                                        filtered_res['LEDGERNAME'] = b.get('LEDGERNAME', None)
                                        filtered_res['AMOUNT'] = b.get('AMOUNT', None)
                                        # print(filtered_res['LEDGERNAME'])
                                        # print(filtered_res['AMOUNT'])
                                        writer.writerow(filtered_res)

                                if "INVENTORYENTRIES.LIST" in res and res["INVENTORYENTRIES.LIST"] is not None:
                                    a = res["INVENTORYENTRIES.LIST"]
                                    # print("c____inventory present")
                                    # print(type(a))

                                    if type(a) == dict:
                                        # print(a)
                                        if "ACCOUNTINGALLOCATIONS.LIST" in a:
                                            filtered_res = dict((k, res[k]) for k in keys)
                                            filtered_res['LEDGERNAME'] = a['ACCOUNTINGALLOCATIONS.LIST'].get(
                                                'LEDGERNAME', None)
                                            filtered_res['AMOUNT'] = a['ACCOUNTINGALLOCATIONS.LIST'].get('AMOUNT', None)
                                            filtered_res['Stock Item'] = a.get('STOCKITEMNAME',None)
                                            filtered_res['Quantity'] = a.get('ACTUALQTY', None)
                                            filtered_res['Rate'] = a.get('RATE', None)
                                            # print(filtered_res['AMOUNT'])
                                            # print(filtered_res['LEDGERNAME'])
                                            writer.writerow(filtered_res)

                                    elif type(a) == list:
                                        # print("here")
                                        for i in a:
                                            # print(i)
                                            if type(i) == dict:
                                                if "ACCOUNTINGALLOCATIONS.LIST" in i:
                                                    filtered_res = dict((k, res[k]) for k in keys)
                                                    filtered_res['LEDGERNAME'] = i['ACCOUNTINGALLOCATIONS.LIST'][
                                                        'LEDGERNAME']
                                                    filtered_res['AMOUNT'] = i['ACCOUNTINGALLOCATIONS.LIST']['AMOUNT']
                                                    filtered_res['Stock Item'] = i.get('STOCKITEMNAME', None)
                                                    filtered_res['Quantity'] = i.get('ACTUALQTY', None)
                                                    filtered_res['Rate'] = i.get('RATE', None)
                                                    #
                                                    writer.writerow(filtered_res)
                                                    # print(filtered_res['LEDGERNAME'])
                                                    # print(filtered_res['AMOUNT'])

                    else:
                        # print("response is dict")
                        if "VOUCHER" in response:
                            res = response['VOUCHER']
                            if "ALLLEDGERENTRIES.LIST" in res and res["ALLLEDGERENTRIES.LIST"] is not None:
                                a = res["ALLLEDGERENTRIES.LIST"]
                                if type(a) == dict:
                                    # print("Dict__________Allledger")
                                    filtered_res = dict((k, res[k]) for k in keys)
                                    filtered_res['LEDGERNAME'] = a.get('LEDGERNAME', None)
                                    filtered_res['AMOUNT'] = a.get('AMOUNT', None)
                                    writer.writerow(filtered_res)
                                    # print(filtered_res['LEDGERNAME'])
                                    # print(filtered_res['AMOUNT'])
                                else:
                                    for i in res['ALLLEDGERENTRIES.LIST']:
                                        filtered_res = dict((k, res[k]) for k in keys)
                                        # print("list__________allledger")
                                        filtered_res['LEDGERNAME'] = i.get('LEDGERNAME', None)
                                        filtered_res['AMOUNT'] = i.get('AMOUNT', None)
                                        writer.writerow(filtered_res)
                                        # print(filtered_res['LEDGERNAME'])
                                        # print(filtered_res['AMOUNT'])

                            if "LEDGERENTRIES.LIST" in res and res["LEDGERENTRIES.LIST"] is not None:
                                b = res["LEDGERENTRIES.LIST"]
                                if type(b) == list:
                                    # print("list_________ledger")
                                    for i in res['LEDGERENTRIES.LIST']:
                                        filtered_res = dict((k, res[k]) for k in keys)
                                        filtered_res['LEDGERNAME'] = i.get('LEDGERNAME', None)
                                        filtered_res['AMOUNT'] = i.get('AMOUNT', None)
                                        # print(filtered_res['LEDGERNAME'])
                                        # print(filtered_res['AMOUNT'])
                                        writer.writerow(filtered_res)
                                elif type(b) == dict:
                                    # print("dict_________ledger")
                                    filtered_res = dict((k, res[k]) for k in keys)
                                    filtered_res['LEDGERNAME'] = b.get('LEDGERNAME', None)
                                    filtered_res['AMOUNT'] = b.get('AMOUNT', None)
                                    # print(filtered_res['LEDGERNAME'])
                                    # print(filtered_res['AMOUNT'])
                                    writer.writerow(filtered_res)

                            if "INVENTORYENTRIES.LIST" in res and res["INVENTORYENTRIES.LIST"] is not None:
                                a = res["INVENTORYENTRIES.LIST"]
                                # print("inventory present 2")

                                if type(a) == dict:
                                    if "ACCOUNTINGALLOCATIONS.LIST" in a:
                                        filtered_res = dict((k, res[k]) for k in keys)
                                        filtered_res['LEDGERNAME'] = a['ACCOUNTINGALLOCATIONS.LIST'].get('LEDGERNAME',
                                                                                                         None)
                                        filtered_res['AMOUNT'] = a['ACCOUNTINGALLOCATIONS.LIST'].get('AMOUNT', None)
                                        # print(filtered_res['AMOUNT'])
                                        # print(filtered_res['LEDGERNAME'])
                                        filtered_res['Stock Item'] = a.get('STOCKITEMNAME', None)
                                        filtered_res['Quantity'] = a.get('ACTUALQTY', None)
                                        filtered_res['Rate'] = a.get('RATE', None)
                                        writer.writerow(filtered_res)

                                elif type(a) == list:
                                    # print("here2")
                                    for i in a:
                                        # print(i)
                                        if type(i) == dict:
                                            if "ACCOUNTINGALLOCATIONS.LIST" in i:
                                                # print("invent_________dict 2")
                                                filtered_res = dict((k, res[k]) for k in keys)
                                                filtered_res['LEDGERNAME'] = i['ACCOUNTINGALLOCATIONS.LIST'][
                                                    'LEDGERNAME']
                                                filtered_res['AMOUNT'] = i['ACCOUNTINGALLOCATIONS.LIST']['AMOUNT']
                                                filtered_res['Stock Item'] = i.get('STOCKITEMNAME', None)
                                                filtered_res['Quantity'] = i.get('ACTUALQTY', None)
                                                filtered_res['Rate'] = i.get('RATE', None)
                                                writer.writerow(filtered_res)
                                                # print(filtered_res['LEDGERNAME'])
                                                # print(filtered_res['AMOUNT'])

                start = start + timedelta(5)
                print(n)
                n += 1
                # break

        print(filename)

    except Exception as e:
        print(e)
