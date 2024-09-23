import openpyxl 
import json


#parameter

catalogo_folder ='CatalogoSSU'
new_folder = catalogo_folder+'\\new'
old_folder = catalogo_folder+'\\old'

unit_test_filename = 'Testbook del catalogo SSU - Unit test_v[#VER].xlsx'
unit_test_new_version = '3.0'
unit_test_old_version = '.2.1'

unit_test_config = {
    'sheet_test': 'Test',
    'sheet_test_column_servizio': 'C',
    'sheet_test_column_request': 'H',
    'sheet_test_column_changed_servizio': 'T'
}


#work unit test sheet test

workbook_new = openpyxl.load_workbook(new_folder+'\\'+unit_test_filename.replace('[#VER]',unit_test_new_version))
workbook_old = openpyxl.load_workbook(old_folder+'\\'+unit_test_filename.replace('[#VER]',unit_test_old_version))

sheet_new = workbook_new[unit_test_config['sheet_test']]
sheet_old = workbook_old[unit_test_config['sheet_test']]

i = 2

while sheet_new[unit_test_config['sheet_test_column_servizio']+str(i)].value:

    #check column servizio
    """
    if (sheet_new[unit_test_config['sheet_test_column_servizio']+str(i)].value!= None):
        text_new_servizio = sheet_new[unit_test_config['sheet_test_column_servizio']+str(i)].value.strip()
    else:
        text_new_servizio = ''

    if (sheet_old[unit_test_config['sheet_test_column_servizio']+str(i)].value!= None):
        text_old_servizio = sheet_old[unit_test_config['sheet_test_column_servizio']+str(i)].value.strip()
    else:
        text_old_servizio = ''

    if (text_new_servizio != text_old_servizio):

        pos_instance_to_into_new = text_new_servizio.find('instance_to_')        
        pos_instance_to_into_old = text_old_servizio.find('instance_to_')

        pos_sended_into_new = text_new_servizio.find('_sended,')
        pos_sended_into_old = text_old_servizio.find('_sended,')

        if (text_new_servizio[:pos_instance_to_into_new]!=text_old_servizio[:pos_instance_to_into_old] or text_new_servizio[pos_sended_into_new:]!=text_old_servizio[pos_sended_into_old:]):
            
            sheet_new[unit_test_config['sheet_test_column_changed_servizio']+str(i)] = text_old_servizio            
    """   
    #check column request    
    try:
        json_new_request = json.loads(sheet_new[unit_test_config['sheet_test_column_request']+str(i)].value.strip().replace('<timestamp01>','0').replace('<start01>','0'))
        
    except ValueError:
        print("  Is valid?: False")
    
    json_old_request = '{\n  "version": 0,\n    "legal_person": "98202900800",\n  "instance_status": [\n    {\n      "state": "started",\n      "timestamp": 0\n    }\n  ],\n   "times": {\n    "start": 0,\n    "max_gg_proc": 60,\n    "max_gg_correction": 5,\n    "max_gg_admissibility": 5,\n    "max_gg_int_resp": 15,\n    "max_gg_int_req": 15,\n    "max_gg_concl_send": 25\n  },\n  "administrative_regime": {\n    "id": "SCIA",\n    "version": "00.00.00"\n  },\n  "usecase_proceedings": [\n    {\n      "code": "USEC-0000413",\n      "version": "00.00.00",\n      "competent_administration": {\n        "ipacode": "IPA_LOC_6",\n        "officecode": "code_LOC_6",\n        "version": "00.00.00",\n        "description": "Ente Comune fittizio locale"\n      },\n       "instance": {\n        "ref": "1",\n        "filename": "Form di default non strutturato",\n        "hash": "5e71d4ffc3a0723b1bcca206fd14c0e217f6d9beb33a5b327ea251d3a54df5af",\n        "alg_hash": "S256",\n        "mime_type": "application/pdf"\n      }\n    }\n  ]\n}'

    '''
    for key in json_new_request:
        print (key+' '+str(key in json_old_request))
    '''
    print(json_new_request)
    i += 1


workbook_new.save(catalogo_folder+'\\'+unit_test_filename.replace('[#VER]',unit_test_new_version+'-'+unit_test_old_version)+"_result.xlsx") 