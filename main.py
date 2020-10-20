import json
import xlrd
import pyexcel
import re
import sys
import os
from datetime import datetime
from shutil import copyfile



# Ignition folder class
class tag_folder:
    def __init__(self, name):
        self.name = name
        self.tagType = 'Folder'
        self.tags = []

# Ignition standard datatype tag class
class tag_standard:
    def __init__(self, name, dataType, plc):
        self.name = name
        self.dataType = dataType
        self.tagType = 'AtomicTag'
        self.tagGroup = 'Direct 1s'
        self.valueSource = 'opc'
        self.opcServer = opc_server
        self.opcItemPath = opc_path_prefix + plc + name
        self.documentation = None

# Ignition UDT tag class
class tag_udt:
    def __init__(self, name, udt_name, plc):
        self.name = name
        self.documentation = None
        self.tagType = 'UdtInstance'
        self.parameters = {
            'req_plc': plc
        }
        # try/except handles cases where desired UDT isn't defined in Ignition
        # Empty 'None' UDT type created for this purpose 
        try:
            self.typeId = udt_dict[udt_name]['udt_ignition']
        except:
            self.typeId = 'None'


# UDT definitions used when constructing tags and searching excel/csv.
# Nested lists used to allow for easier future automation/GUI input
# Primary list:     one UDT (secondary list) per element
# Secondary list:   [UDT Name, Ignition UDT Path, PLC UDT, PLC UDT Alias (for array-aliased tags - optional), PLC UDT name prefix]
# PLC UDT generally refers to HMIData tag UDT.
# PLC UDT array primarily used for AI/DI imports, where tags are largely defined as array aliases
udt_raw = [
    ['CMCSM',   'Standard/CMCSM/gtypCMCSM',     'gtypCMCSMHmiData',     None,               'gt_CMCSMHmiData_'],
    ['CMVSM',   'Standard/CMVSM/gtypCMVSM',     'gtypCMVSMHmiData',     None,               'gt_CMVSMHmiData_'],
    ['CMIV',    'Standard/CMIV/gtypCMIV',       'gtypCMIVHmiData',      None,               'gt_CMIVHmiData_'],
    ['CMCV',    'Standard/CMCV/gtypCMCV',       'gtypCMCVHmiData',      None,               'gt_CMCVHmiData_'],
    ['CMDD',    'Standard/CMDD/gtypCMDD',       'gtypCMDDHmiData',      None,               'gt_CMDDHmiData_'],
    ['CMID',    'Standard/CMID/gtypCMID',       'gtypCMIDHmiData',      None,               'gt_CMIDHmiData_'],
    ['CMCD',    'Standard/CMCD/gtypCMCD',       'gtypCMCDHmiData',      None,               'gt_CMCDHmiData_'],
    ['PIDv1',   'Standard/PID/gtypPIDv1',       'gtypPIDHmiData',       None,               'gt_PidHmiData_'],
    ['PIDv2',   'Standard/PID/gtypPIDv2',       'gtypPID_v2_HmiData',   None,               'gt_PidHmiData_'],
    ['AI',      'Standard/AI/gtypAI',           'gtypAIHmiData',        'gtaAIHmiData',     'gt_HmiData_'],
    ['DI',      'Standard/DI/gtypDI',           'gtypDIHmiData',        'gtaDIHmiData',     'gt_HmiData_']   
]

# UDT dictionary enables simple loops to create tags lists iteratively. Example dict entry:
# {   
#     'CMCSM': {
#         'udt_ignition': 'Standard/CMCSM/gtypCMCSM',
#         'udt_plc': 'gtypCMCSMHmiData',
#         'udt_plc_alias': None,
#         'udt_plc_prefix': 'gt_CMCSMHmiData_'
#     },
#     ...
# }
udt_dict = {}
for i in range(len(udt_raw)):
    temp = {}
    temp['udt_ignition']    = udt_raw[i][1]
    temp['udt_plc']         = udt_raw[i][2]
    temp['udt_plc_alias']   = udt_raw[i][3]
    temp['udt_plc_prefix']  = udt_raw[i][4]

    udt_dict[udt_raw[i][0]] = temp


# Standard tag definitions used when constructing tags and searching excel/csv.
# Nested lists used to allow for easier future automation/GUI input
# Primary list:     one standard tag type (secondary list) per element
# Secondary list:   [PLC datatype, Ignition datatype, Ignition equivalent array datatype (if one exists)]
standard_raw = [
    ['BOOL',    'Boolean',  'Boolean Array'],
    ['INT',     'Short',    'Short Array'],
    ['DINT',    'Integer',  'Integer Array'],
    ['REAL',    'Float',    'Float Array'],
    ['STRING',  'String',   'String Array'],
]

# Standard tag dictionary enables simple loops to create tags lists iteratively. Example dict entry:
# {   
#     'BOOL': {
#         'ignition_datatype': 'Boolean',
#         'ignition_datatype_array': 'Boolean Array'
#     },
#     ...
# }
standard_dict = {}
for i in range(len(standard_raw)):
    temp = {}
    temp['ignition_datatype']       = standard_raw[i][1]
    temp['ignition_datatype_array'] = standard_raw[i][2]

    standard_dict[standard_raw[i][0]] = temp


# Load excel/csv workbook.
# If CSV, convert to excel book, save it as a temp file
# If excel, copy to temp file to prevent read/file open errors
workbook_path = 'ignition-tag-import/tags.xlsx'
temp_workbook_path = 'ignition-tag-import/temp/temp_workbook.xlsx'
if workbook_path.lower().endswith('.csv'):
    csv_workbook = pyexcel.get_sheet(file_name=workbook_path, delimiter=',')
    csv_workbook.save_as(temp_workbook_path)
else:
    copyfile(workbook_path, temp_workbook_path)
workbook = xlrd.open_workbook(temp_workbook_path)
sheet = workbook.sheet_by_index(0)


# Define plc shortcut, opc details
# Ignition OPC UA server required additional prefix for OPC item path
plc = '[B8_CP_001]'
opc_server = 'RSLinx'
if 'Ignition OPC UA Server' in opc_server:
    opc_path_prefix = 'ns=1;s='
else:
    opc_path_prefix = ''


# Define import tag folder
import_folder = tag_folder('Imported (Delete before startup)')

# Loop once per UDT
for udt_name in udt_dict:
    temp_folder = tag_folder(udt_name)
    
    # Check each row in sheet
    for row in range(sheet.nrows):
        # Unaliased UDT build
        # Tag must be 'TAG' type, global scope, and match desired PLC datatype
        if sheet.cell_value(row,0) == 'TAG' and sheet.cell_value(row,1) == '' and sheet.cell_value(row,4) == udt_dict[udt_name]['udt_plc']:
            temp_name = sheet.cell_value(row,2)
            temp_name = re.sub(udt_dict[udt_name]['udt_plc_prefix'], '', temp_name, flags=re.I)
            temp_udt = tag_udt(temp_name, udt_name, plc)
            temp_udt.documentation = sheet.cell_value(row,3)
            temp_folder.tags.append(temp_udt.__dict__) 
        
        # Aliased UDT build (if alias defined)
        elif udt_dict[udt_name]['udt_plc_alias'] is not None:
            # Tag must be 'ALIAS' type and be derived from desired PLC datatype
            if sheet.cell_value(row,0) == 'ALIAS' and str(sheet.cell_value(row,5)).startswith(udt_dict[udt_name]['udt_plc_alias']):
                temp_name = sheet.cell_value(row,2)
                temp_name = re.sub(udt_dict[udt_name]['udt_plc_prefix'], '', temp_name, flags=re.I)
                temp_udt = tag_udt(temp_name, udt_name, plc)
                temp_udt.documentation = sheet.cell_value(row,3)
                temp_folder.tags.append(temp_udt.__dict__)
    
    # After creating all tags of a given udt, append UDT folder to import folder
    import_folder.tags.append(temp_folder.__dict__)


# Define standard tag folder
standard_folder = tag_folder('Standard')

# Loop once per standard type
for datatype in standard_dict:
    temp_folder = tag_folder(datatype)

    # Check each row in sheet
    for row in range(sheet.nrows):
        # Standalone tag build
        # Tag must be 'TAG' type, global scope, contain 'hmidata' tagname, and match desired PLC datatype
        if sheet.cell_value(row,0) == 'TAG' and sheet.cell_value(row,1) == '' and ('hmidata' in str(sheet.cell_value(row,2)).lower()) and str(sheet.cell_value(row,4)).startswith(datatype):
            temp_name = sheet.cell_value(row,2)
            
            # determine datatype based on whether tag is standalone (e.g. BOOL) or arrayed (e.g. BOOL[10])
            if sheet.cell_value(row,4) == datatype:
                temp_standard = tag_standard(temp_name, standard_dict[datatype]['ignition_datatype'], plc)
            elif str(sheet.cell_value(row,4)).startswith(datatype + '[') and standard_dict[datatype]['ignition_datatype_array'] is not None:
                temp_standard = tag_standard(temp_name, standard_dict[datatype]['ignition_datatype_array'], plc)
            else:
                continue
            temp_standard.documentation = sheet.cell_value(row,3)
            temp_folder.tags.append(temp_standard.__dict__)

    # After creating all tags of a given standard datatype, append datatype folder to standard folder
    standard_folder.tags.append(temp_folder.__dict__)

# Append standard tags to import tag folder
import_folder.tags.append(standard_folder.__dict__)


# Create tag import JSON file
# output_path example: 'ignition-tag-import/output/[B8_CP_001] tag import 03Jun2020 132501.json'
datetime_string = datetime.now().strftime('%d%b%Y %H%M%S')
output_path = 'ignition-tag-import/output/' + plc + ' tag import ' + datetime_string + '.json'
output_file = open(output_path, 'w+')

# Dump import_folder to file with indentation and alphabetical sorting
json.dump(import_folder.__dict__, output_file, indent=4, sort_keys=True)

# Delete temp excel file after script completes
os.remove(temp_workbook_path)