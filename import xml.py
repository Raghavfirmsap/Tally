import xml.etree.ElementTree as ET
import xmltodict
import json

tree = ET.parse('data-out.xml')
xml_data = tree.getroot()
xmlstr = ET.tostring(xml_data, encoding='utf-8', method='xml')

data_dict = dict(xmltodict.parse(xmlstr))

print(data)
