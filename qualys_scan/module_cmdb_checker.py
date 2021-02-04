# To add a new cell, type '# %%'
# To add a new markdown cell, type '# %% [markdown]'
# %%
import time, os, requests,shutil
import pandas as pd
columns = shutil.get_terminal_size().columns


# %%
def connection_check():
    API = '0a66e9aecdc8c49191309691d8c340b920a060b3'
    endpoint = 'https://cmdb.netledger.com/api/dc-hosts/'
    headers = {'Authorization': 'Token '+ API}
    response = requests.get(endpoint, headers=headers)
    return (response)


# %%
def cmdb_search (sn):
    API = '0a66e9aecdc8c49191309691d8c340b920a060b3'
    endpoint = f'https://cmdb.netledger.com/api/data-center-assets/?sn={sn}'
#     url = endpoint.format(dns=dns)
    headers = {'Authorization': 'Token '+ API}
    response = requests.get(endpoint, headers=headers)
    result = response.json()
    return result

def cmdb_id_search(id):
    API = '0a66e9aecdc8c49191309691d8c340b920a060b3'
    endpoint = f'https://cmdb.netledger.com/api/data-center-assets/{id}'
#     url = endpoint.format(dns=dns)
    headers = {'Authorization': 'Token '+ API}
    response = requests.get(endpoint, headers=headers)
    result = response.json()
    return result


# %%
print("*".center(columns,'*'))       
print ("This program will check the Server SN and AT are in CMDB.".center(columns))
print ("This program will also check if the module is in the correct parent server.".center(columns))
print("Make sure that you are connected to the VPN.".center(columns))
print("*".center(columns,'*'))   


# %%
print("Checking VPN Connection")
status = connection_check().status_code
if status == 200:
     print('[OK]')
else:
    while status != 200:
        print('Check VPN connection before pressing "Enter": ')
        x = input('')
        status = connection_check().status_code


# %%
#Obtaining File
servers_file = input("enter server list file name location ex: ~/Downloads/servers_file.xlsx: ")
pcp_check = os.path.isfile(servers_file)
while servers_file == False:
    servers_file = input("enter server list file name location ex: ~/Downloads/servers_file.xlsx: ")
    pcp_check = os.path.isfile(servers_file)
print(f"{servers_file} file found: " + str(pcp_check))


# %%
workbook = pd.ExcelFile(f'{servers_file}')


# %%
data = pd.read_excel(workbook)


# %%
def chassis_check(sn,at):
    with open("output.txt", "a") as f:
        x = cmdb_search (sn)
        if sn == x['results'][0]['sn']:
            print(f"Parent Serial: {sn} matches in CMDB", file=f)
        else:
            print(f"!!! Parent Serial: {sn} doesn't match in CMDB !!!", file=f)
        if at == x['results'][0]['barcode']:
            print(f"Parent Asset Tag: {at} matches in CMDB", file=f)
        else:
            print(f"!!! Parent Asset Tag: {at} doesn't match in CMDB !!!", file=f)

        if "X7-2c Chassis" == x['results'][0]['model']['name']:
            print(f"Parent Model Type:X7-2c Chassis matches in CMDB", file=f)
        else:
            print(f"!!! Parent Model Type:X7-2c Chassis doesn't match in CMDB !!!", file=f)
        print('\n', file=f)
        
def strp(x):
    x = str(x).strip()
    return x


# %%
def mod_check(slot_no,msn1,mat1,psn):
    with open("output.txt", "a") as f:
        x = cmdb_search (msn1)
        if msn1 == x['results'][0]['sn']:
            print(f"Module {slot_no} Serial: {msn1} matches in CMDB", file=f)
        else:
            print(f"!!! Module {slot_no} Serial: {msn1} doesn't match in CMDB !!!", file=f)
        if mat1 == x['results'][0]['barcode']:
            print(f"Module {slot_no} Asset Tag: {mat1} matches in CMDB", file=f)
        else:
            print(f"!!! Module {slot_no} Asset Tag: {mat1} doesn't match in CMDB !!!", file=f)

        if "X7-2c Module" == x['results'][0]['model']['name']:
            print(f"Module {slot_no} Model Type:X7-2c Module matches in CMDB", file=f)
        elif "X7-2c Module " == x['results'][0]['model']['name']:
            print(f"Module {slot_no} Model Type:X7-2c Module matches in CMDB", file=f)
        else:
            print(f"!!! Module {slot_no} Model Type:X7-2c Module doesn't match in CMDB !!!", file=f)

        if '1' == x['results'][0]['slot_no']:
            print(f"Module {slot_no} Slot {slot_no} matches in CMDB", file=f)
        else:
            print(f"!!! Module {slot_no} Slot {slot_no} doesn't match in CMDB !!!", file=f)

        pid =  cmdb_id_search (x['results'][0]['parent']['id'])
        if psn ==pid['sn']:
            print(f"Module {slot_no} SN: {msn1} / Parent SN:{psn} matches CMDB", file=f)
        else:
            print(f"!!! Module {slot_no} SN: {msn1}/Parent SN:{psn} Doesn't match CMDB !!!", file=f)
        print('\n', file=f)
    


# %%



# %%
for psn,pat,msn1,mat1,msn2,mat2 in zip(data['Parent SN'],data['Parent AT'],data['Module 1 SN'],data['Module 1 AT'],data['Module 2 SN'],data['Module 2 AT']):
    psn = strp(psn)
    pat = strp(pat)
    msn1 = strp(msn1)
    mat1 = strp(mat1)
    msn2 =strp(msn2)
    mat2 = strp(mat2)
    print(f"Checking Chasis : {psn}")
    chassis_check(psn,pat)
    print(f"Checking Module 1 : {msn1} ")
    mod_check(1,msn1,mat1,psn)
    print(f"Checking Module 2 : {msn2}")
    mod_check(2,msn1,mat1,psn)
    
print("\n Output.txt Exported")


# %%



# %%



# %%



# %%



