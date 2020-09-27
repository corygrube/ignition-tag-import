import json
import xlrd
import re

# Open/create tag import JSON file
# Load excel book
# Load UDT dicts from json files
json_file = open('tag import.json', 'w+')
book = xlrd.open_workbook('tags.xlsx')
AI_json = open('json templates/AI.json', 'r')
AI_temp = json.load(AI_json)
AI_tags = []
DI_json = open('json templates/DI.json', 'r')
DI_temp = json.load(DI_json)
DI_tags = []
CM2SM_json = open('json templates/CM2SM.json', 'r')
CM2SM_temp = json.load(CM2SM_json)
CM2SM_tags = []
CMCSM_json = open('json templates/CMCSM.json', 'r')
CMCSM_temp = json.load(CMCSM_json)
CMCSM_tags = []
CMVSM_json = open('json templates/CMVSM.json', 'r')
CMVSM_temp = json.load(CMVSM_json)
CMVSM_tags = []
CMIV_json = open('json templates/CMIV.json', 'r')
CMIV_temp = json.load(CMIV_json)
CMIV_tags = []
CMCV_json = open('json templates/CMCV.json', 'r')
CMCV_temp = json.load(CMCV_json)
CMCV_tags = []
CMDD_json = open('json templates/CMDD.json', 'r')
CMDD_temp = json.load(CMDD_json)
CMDD_tags = []
CMID_json = open('json templates/CMID.json', 'r')
CMID_temp = json.load(CMID_json)
CMID_tags = []
CMCD_json = open('json templates/CMCD.json', 'r')
CMCD_temp = json.load(CMCD_json)
CMCD_tags = []
PIDv1_json = open('json templates/PIDv1.json', 'r')
PIDv1_temp = json.load(PIDv1_json)
PIDv1_tags = []
PIDv2_json = open('json templates/PIDv2.json', 'r')
PIDv2_temp = json.load(PIDv2_json)
PIDv2_tags = []


# Set PLC name, create import dict
plc = '[B8_CP_001]'
import_dict = {}

# Load sheet. CSV, so only 1 sheet (index 0)
sheet = book.sheet_by_index(0)

# Read tag addresses/descriptions from excel, load into tag dict
for row in range(sheet.nrows):
    # Evaluate TAG rows (not alias, comment etc).
    if sheet.cell_value(row,0) == 'TAG':
        if sheet.cell_value(row,4) == 'gtypCM2SMHmiData' and sheet.cell_value(row,1) == '':
            name = str(sheet.cell_value(row,2))
            CM2SM_temp['name'] = re.sub('gt_CM2SMHmiData_', '', name, flags=re.I)
            CM2SM_temp['documentation'] = str(sheet.cell_value(row,3))
            CM2SM_temp['parameters']['req_plc'] = plc
            CM2SM_tags.append(CM2SM_temp.copy())
        
        elif sheet.cell_value(row,4) == 'gtypCMCSMHmiData' and sheet.cell_value(row,1) == '':
            name = str(sheet.cell_value(row,2))
            CMCSM_temp['name'] = re.sub('gt_CMCSMHmiData_', '', name, flags=re.I)
            CMCSM_temp['documentation'] = str(sheet.cell_value(row,3))
            CMCSM_temp['parameters']['req_plc'] = plc
            CMCSM_tags.append(CMCSM_temp.copy())
            
        elif sheet.cell_value(row,4) == 'gtypCMVSMHmiData' and sheet.cell_value(row,1) == '':
            name = str(sheet.cell_value(row,2))
            CMVSM_temp['name'] = re.sub('gt_CMVSMHmiData_', '', name, flags=re.I)
            CMVSM_temp['documentation'] = str(sheet.cell_value(row,3))
            CMVSM_temp['parameters']['req_plc'] = plc
            CMVSM_tags.append(CMVSM_temp.copy())
        
        elif sheet.cell_value(row,4) == 'gtypCMIVHmiData' and sheet.cell_value(row,1) == '':
            name = str(sheet.cell_value(row,2))
            CMIV_temp['name'] = re.sub('gt_CMIVHmiData_', '', name, flags=re.I)
            CMIV_temp['documentation'] = str(sheet.cell_value(row,3))
            CMIV_temp['parameters']['req_plc'] = plc
            CMIV_tags.append(CMIV_temp.copy())

        elif sheet.cell_value(row,4) == 'gtypCMCVHmiData' and sheet.cell_value(row,1) == '':
            name = str(sheet.cell_value(row,2))
            CMCV_temp['name'] = re.sub('gt_CMCVHmiData_', '', name, flags=re.I)
            CMCV_temp['documentation'] = str(sheet.cell_value(row,3))
            CMCV_temp['parameters']['req_plc'] = plc
            CMCV_tags.append(CMCV_temp.copy())

        elif sheet.cell_value(row,4) == 'gtypCMDDHmiData' and sheet.cell_value(row,1) == '':
            name = str(sheet.cell_value(row,2))
            CMDD_temp['name'] = re.sub('gt_CMDDHmiData_', '', name, flags=re.I)
            CMDD_temp['documentation'] = str(sheet.cell_value(row,3))
            CMDD_temp['parameters']['req_plc'] = plc
            CMDD_tags.append(CMDD_temp.copy())
        
        elif sheet.cell_value(row,4) == 'gtypCMIDHmiData' and sheet.cell_value(row,1) == '':
            name = str(sheet.cell_value(row,2))
            CMID_temp['name'] = re.sub('gt_CMIDHmiData_', '', name, flags=re.I)
            CMID_temp['documentation'] = str(sheet.cell_value(row,3))
            CMID_temp['parameters']['req_plc'] = plc
            CMID_tags.append(CMID_temp.copy())
        
        elif sheet.cell_value(row,4) == 'gtypCMCDHmiData' and sheet.cell_value(row,1) == '':
            name = str(sheet.cell_value(row,2))
            CMCD_temp['name'] = re.sub('gt_CMCDHmiData_', '', name, flags=re.I)
            CMCD_temp['documentation'] = str(sheet.cell_value(row,3))
            CMCD_temp['parameters']['req_plc'] = plc
            CMCD_tags.append(CMCD_temp.copy())

        elif sheet.cell_value(row,4) == 'gtypPidHmiData' and sheet.cell_value(row,1) == '':
            name = str(sheet.cell_value(row,2))
            PIDv1_temp['name'] = re.sub('gt_PidHmiData_', '', name, flags=re.I)
            PIDv1_temp['documentation'] = str(sheet.cell_value(row,3))
            PIDv1_temp['parameters']['req_plc'] = plc
            PIDv1_tags.append(PIDv1_temp.copy())
            
        elif sheet.cell_value(row,4) == 'gtypPid_V2_HmiData' and sheet.cell_value(row,1) == '':
            name = str(sheet.cell_value(row,2))
            PIDv2_temp['name'] = re.sub('gt_PidHmiData_', '', name, flags=re.I)
            PIDv2_temp['documentation'] = str(sheet.cell_value(row,3))
            PIDv2_temp['parameters']['req_plc'] = plc
            PIDv2_tags.append(PIDv2_temp.copy())
    
    # Evaluate ALIAS rows (array-indexed IO, etc).
    elif sheet.cell_value(row,0) == 'ALIAS':
        if str(sheet.cell_value(row,5)).startswith('gtaAIHmiData'):
            name = str(sheet.cell_value(row,2))
            AI_temp['name'] = re.sub('gt_HmiData_', '', name, flags=re.I)
            AI_temp['documentation'] = str(sheet.cell_value(row,3))
            AI_temp['parameters']['req_plc'] = plc
            AI_tags.append(AI_temp.copy())
        
        elif str(sheet.cell_value(row,5)).startswith('gtaDIHmiData'):
            name = str(sheet.cell_value(row,2))
            DI_temp['name'] = re.sub('gt_HmiData_', '', name, flags=re.I)
            DI_temp['documentation'] = str(sheet.cell_value(row,3))
            DI_temp['parameters']['req_plc'] = plc
            DI_tags.append(DI_temp.copy())

# Create final dictionary with all identified tags
import_dict['tags'] = [{},{},{},{},{},{},{},{},{},{},{}]
import_dict['tags'][0]['name'] = 'CMCSM'
import_dict['tags'][0]['tagType'] = 'Folder'
import_dict['tags'][0]['tags'] = CMCSM_tags
import_dict['tags'][1]['name'] = 'CMVSM'
import_dict['tags'][1]['tagType'] = 'Folder'
import_dict['tags'][1]['tags'] = CMVSM_tags
import_dict['tags'][2]['name'] = 'CMIV'
import_dict['tags'][2]['tagType'] = 'Folder'
import_dict['tags'][2]['tags'] = CMIV_tags
import_dict['tags'][3]['name'] = 'CMCV'
import_dict['tags'][3]['tagType'] = 'Folder'
import_dict['tags'][3]['tags'] = CMCV_tags
import_dict['tags'][4]['name'] = 'CMDD'
import_dict['tags'][4]['tagType'] = 'Folder'
import_dict['tags'][4]['tags'] = CMDD_tags
import_dict['tags'][5]['name'] = 'CMID'
import_dict['tags'][5]['tagType'] = 'Folder'
import_dict['tags'][5]['tags'] = CMID_tags
import_dict['tags'][6]['name'] = 'CMCD'
import_dict['tags'][6]['tagType'] = 'Folder'
import_dict['tags'][6]['tags'] = CMCD_tags
import_dict['tags'][7]['name'] = 'PID'
import_dict['tags'][7]['tagType'] = 'Folder'
import_dict['tags'][7]['tags'] = PIDv1_tags
# if PIDv2_tags != []: 
#     import_dict['tags'][7]['tags'].append(PIDv2_tags)         # Think this needs a for loop instead of a single append

# for key in PIDv2_tags:
#     import_dict['tags'][7]['tags'].append(PIDv2_tags[key])    # 1st attempt. no good as is. Think it won't work with blank list either.
import_dict['tags'][8]['name'] = 'AI'
import_dict['tags'][8]['tagType'] = 'Folder'
import_dict['tags'][8]['tags'] = AI_tags
import_dict['tags'][9]['name'] = 'DI'
import_dict['tags'][9]['tagType'] = 'Folder'
import_dict['tags'][9]['tags'] = DI_tags
import_dict['tags'][10]['name'] = 'CM2SM'
import_dict['tags'][10]['tagType'] = 'Folder'
import_dict['tags'][10]['tags'] = CM2SM_tags


# Dump import_dict to file
json.dump(import_dict, json_file)

# udt_dict = {
#     'CMCSM': 'Control Modules/CMCSM/gtypCMCSM',
#     'CMVSM': 'Control Modules/CMVSM/gtypCMVSM',
#     'CMIV': 'Control Modules/CMIV/gtypCMIV',
#     'CMCV': 'Control Modules/CMCV/gtypCMCV',
#     'CMDD': 'Control Modules/CMDD/gtypCMDD',
#     'CMID': 'Control Modules/CMID/gtypCMID',
#     'CMCD': 'Control Modules/CMCD/gtypCMCD',
#     'PIDv1': 'Control Modules/PID/gtypPIDv1',
#     'PIDv2': 'Control Modules/PID/gtypPIDv2',
#     'AI': 'Control Modules/AI/gtypAI',
#     'DI': 'Control Modules/DI/gtypDI',
# }

# class folder:
#     def __init__(self, name):
#         self.name = name
#         self.tagType = 'Folder'
#         self.tags = []

# class tag:
#     def __init__(self, name, dataType):
#         self.name = name
#         self.dataType = dataType
#         self.tagType = 'AtomicTag'
#         self.tagGroup = 'Direct 1s'
#         self.valueSource = 'opc'
#         self.opcServer = 'RSLinx'
#         self.opcItemPath = plc + name
#         self.documentation = None

# class udt:
#     def __init__(self, name, udtType):
#         self.name = name
#         self.documentation = None
#         self.tagType = 'UdtInstance'
#         self.parameters = {
#             'req_plc': plc
#         }

#         try:
#             self.typeId = udt_dict[udtType]
#         except:
#             self.typeId = 'None'