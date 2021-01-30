import pandas as pd
import shutil,time, os.path,requests

import concurrent.futures
pd.options.mode.chained_assignment = None
columns = shutil.get_terminal_size().columns
import requests
def connection_check():
    API = '0a66e9aecdc8c49191309691d8c340b920a060b3'
    endpoint = 'https://cmdb.netledger.com/api/dc-hosts/'
    headers = {'Authorization': 'Token '+ API}
    response = requests.get(endpoint, headers=headers)
    return response
def main_ip_search (ip):
    API = '0a66e9aecdc8c49191309691d8c340b920a060b3'
    ip = ip
    endpoint = 'https://cmdb.netledger.com/api/ipaddresses/?address={ip}'
    url = endpoint.format(ip=ip)
    headers = {'Authorization': 'Token '+ API}
    response = requests.get(url, headers=headers)
    result = response.json()
    count = result['count']
    if count == 1:
        print(str(ip) + " IP IN CMDB ")
        workbook.main['IP in CMDB'][workbook.main['IP'] == ip] = 'IP in CMDB'
    elif count < 1:
        print(str(ip) + " IP NOT in CMDB ")

        workbook.main['IP in CMDB'][workbook.main['IP'] == ip] = 'IP NOT in CMDB'
    elif count > 1:
        print(str(ip) + " IP IN CMDB + ")

def main_dns_search (dns):
    API = '0a66e9aecdc8c49191309691d8c340b920a060b3'
    dns = dns
    endpoint = 'https://cmdb.netledger.com/api/dc-hosts/?hostname={dns}'
    url = endpoint.format(dns=dns)
    headers = {'Authorization': 'Token '+ API}
    response = requests.get(url, headers=headers)
    result = response.json()
    count = result['count']
    if count == 1:
        workbook.main['DNS in CMDB'][workbook.main['DNS'] == dns] = 'DNS in CMDB'
        print(str(dns)," DNS in cmdb")
    elif count < 1:
        workbook.main['DNS in CMDB'][workbook.main['DNS'] == dns] = 'DNS NOT in CMDB'
        print(str(dns)," DNS NOT in cmdb")
    elif count > 1:
        workbook.main['DNS in CMDB'][workbook.main['DNS'] == dns] = 'DNS in CMDB + '
        print(str(dns)," DNS in CMDB +")  
        
def vpn_ip_search (ip):
    API = '0a66e9aecdc8c49191309691d8c340b920a060b3'
    ip = ip
    endpoint = 'https://cmdb.netledger.com/api/ipaddresses/?address={ip}'
    url = endpoint.format(ip=ip)
    headers = {'Authorization': 'Token '+ API}
    response = requests.get(url, headers=headers)
    result = response.json()
    count = result['count']
    if count == 1:
        print(str(ip) + " IP IN CMDB ")
        workbook.vpn['IP in CMDB'][workbook.vpn['IP'] == ip] = 'IP in CMDB'
    elif count < 1:
        print(str(ip) + " IP NOT in CMDB ")

        workbook.vpn['IP in CMDB'][workbook.vpn['IP'] == ip] = 'IP NOT in CMDB'
    elif count > 1:
        print(str(ip) + " IP IN CMDB + ")

def vpn_dns_search (dns):
    API = '0a66e9aecdc8c49191309691d8c340b920a060b3'
    dns = dns
    endpoint = 'https://cmdb.netledger.com/api/dc-hosts/?hostname={dns}'
    url = endpoint.format(dns=dns)
    headers = {'Authorization': 'Token '+ API}
    response = requests.get(url, headers=headers)
    result = response.json()
    count = result['count']
    if count == 1:
        workbook.vpn['DNS in CMDB'][workbook.vpn['DNS'] == dns] = 'DNS in CMDB'
        print(str(dns)," DNS in cmdb")
    elif count < 1:
        workbook.vpn['DNS in CMDB'][workbook.vpn['DNS'] == dns] = 'DNS NOT in CMDB'
        print(str(dns)," DNS NOT in cmdb")
    elif count > 1:
        workbook.vpn['DNS in CMDB'][workbook.vpn['DNS'] == dns] = 'DNS in CMDB + '
        print(str(dns)," DNS in CMDB +")  
        
def none_ip_search (ip):
    API = '0a66e9aecdc8c49191309691d8c340b920a060b3'
    ip = ip
    endpoint = 'https://cmdb.netledger.com/api/ipaddresses/?address={ip}'
    url = endpoint.format(ip=ip)
    headers = {'Authorization': 'Token '+ API}
    response = requests.get(url, headers=headers)
    result = response.json()
    count = result['count']
    if count == 1:
        print(str(ip) + " IP IN CMDB ")
        workbook.none['IP in CMDB'][workbook.none['IP'] == ip] = 'IP in CMDB'
    elif count < 1:
        print(str(ip) + " IP NOT in CMDB ")

        workbook.none['IP in CMDB'][workbook.none['IP'] == ip] = 'IP NOT in CMDB'
    elif count > 1:
        print(str(ip) + " IP IN CMDB + ")

def none_dns_search (dns):
    API = '0a66e9aecdc8c49191309691d8c340b920a060b3'
    dns = dns
    endpoint = 'https://cmdb.netledger.com/api/dc-hosts/?hostname={dns}'
    url = endpoint.format(dns=dns)
    headers = {'Authorization': 'Token '+ API}
    response = requests.get(url, headers=headers)
    result = response.json()
    count = result['count']
    if count == 1:
        workbook.none['DNS in CMDB'][workbook.none['DNS'] == dns] = 'DNS in CMDB'
        print(str(dns)," DNS in cmdb")
    elif count < 1:
        workbook.none['DNS in CMDB'][workbook.none['DNS'] == dns] = 'DNS NOT in CMDB'
        print(str(dns)," DNS NOT in cmdb")
    elif count > 1:
        workbook.none['DNS in CMDB'][workbook.none['DNS'] == dns] = 'DNS in CMDB + '
        print(str(dns)," DNS in CMDB +")  

def console_ip_search (ip):
    API = '0a66e9aecdc8c49191309691d8c340b920a060b3'
    ip = ip
    endpoint = 'https://cmdb.netledger.com/api/ipaddresses/?address={ip}'
    url = endpoint.format(ip=ip)
    headers = {'Authorization': 'Token '+ API}
    response = requests.get(url, headers=headers)
    result = response.json()
    count = result['count']
    if count == 1:
        print(str(ip) + " IP IN CMDB ")
        workbook.console['IP in CMDB'][workbook.console['IP'] == ip] = 'IP in CMDB'
    elif count < 1:
        print(str(ip) + " IP NOT in CMDB ")

        workbook.console['IP in CMDB'][workbook.console['IP'] == ip] = 'IP NOT in CMDB'
    elif count > 1:
        print(str(ip) + " IP IN CMDB + ")

def console_dns_search (dns):
    API = '0a66e9aecdc8c49191309691d8c340b920a060b3'
    dns = dns
    endpoint = 'https://cmdb.netledger.com/api/dc-hosts/?hostname={dns}'
    url = endpoint.format(dns=dns)
    headers = {'Authorization': 'Token '+ API}
    response = requests.get(url, headers=headers)
    result = response.json()
    count = result['count']
    if count == 1:
        workbook.console['DNS in CMDB'][workbook.console['DNS'] == dns] = 'DNS in CMDB'
        print(str(dns)," DNS in cmdb")
    elif count < 1:
        workbook.console['DNS in CMDB'][workbook.console['DNS'] == dns] = 'DNS NOT in CMDB'
        print(str(dns)," DNS NOT in cmdb")
    elif count > 1:
        workbook.console['DNS in CMDB'][workbook.console['DNS'] == dns] = 'DNS in CMDB + '
        print(str(dns)," DNS in CMDB +")  
        
def interface_ip_search (ip):
    API = '0a66e9aecdc8c49191309691d8c340b920a060b3'
    ip = ip
    endpoint = 'https://cmdb.netledger.com/api/ipaddresses/?address={ip}'
    url = endpoint.format(ip=ip)
    headers = {'Authorization': 'Token '+ API}
    response = requests.get(url, headers=headers)
    result = response.json()
    count = result['count']
    if count == 1:
        print(str(ip) + " IP IN CMDB ")
        workbook.interface['IP in CMDB'][workbook.interface['IP'] == ip] = 'IP in CMDB'
    elif count < 1:
        print(str(ip) + " IP NOT in CMDB ")

        workbook.interface['IP in CMDB'][workbook.interface['IP'] == ip] = 'IP NOT in CMDB'
    elif count > 1:
        print(str(ip) + " IP IN CMDB + ")

def interface_dns_search (dns):
    API = '0a66e9aecdc8c49191309691d8c340b920a060b3'
    dns = dns
    endpoint = 'https://cmdb.netledger.com/api/dc-hosts/?hostname={dns}'
    url = endpoint.format(dns=dns)
    headers = {'Authorization': 'Token '+ API}
    response = requests.get(url, headers=headers)
    result = response.json()
    count = result['count']
    if count == 1:
        workbook.interface['DNS in CMDB'][workbook.interface['DNS'] == dns] = 'DNS in CMDB'
        print(str(dns)," DNS in cmdb")
    elif count < 1:
        workbook.interface['DNS in CMDB'][workbook.interface['DNS'] == dns] = 'DNS NOT in CMDB'
        print(str(dns)," DNS NOT in cmdb")
    elif count > 1:
        workbook.interface['DNS in CMDB'][workbook.interface['DNS'] == dns] = 'DNS in CMDB + '
        print(str(dns)," DNS in CMDB +")  
        
def nouse_ip_search (ip):
    API = '0a66e9aecdc8c49191309691d8c340b920a060b3'
    ip = ip
    endpoint = 'https://cmdb.netledger.com/api/ipaddresses/?address={ip}'
    url = endpoint.format(ip=ip)
    headers = {'Authorization': 'Token '+ API}
    response = requests.get(url, headers=headers)
    result = response.json()
    count = result['count']
    if count == 1:
        print(str(ip) + " IP IN CMDB ")
        workbook.nouse['IP in CMDB'][workbook.nouse['IP'] == ip] = 'IP in CMDB'
    elif count < 1:
        print(str(ip) + " IP NOT in CMDB ")

        workbook.nouse['IP in CMDB'][workbook.nouse['IP'] == ip] = 'IP NOT in CMDB'
    elif count > 1:
        print(str(ip) + " IP IN CMDB + ")

def nouse_dns_search (dns):
    API = '0a66e9aecdc8c49191309691d8c340b920a060b3'
    dns = dns
    endpoint = 'https://cmdb.netledger.com/api/dc-hosts/?hostname={dns}'
    url = endpoint.format(dns=dns)
    headers = {'Authorization': 'Token '+ API}
    response = requests.get(url, headers=headers)
    result = response.json()
    count = result['count']
    if count == 1:
        workbook.nouse['DNS in CMDB'][workbook.nouse['DNS'] == dns] = 'DNS in CMDB'
        print(str(dns)," DNS in cmdb")
    elif count < 1:
        workbook.nouse['DNS in CMDB'][workbook.nouse['DNS'] == dns] = 'DNS NOT in CMDB'
        print(str(dns)," DNS NOT in cmdb")
    elif count > 1:
        workbook.nouse['DNS in CMDB'][workbook.nouse['DNS'] == dns] = 'DNS in CMDB + '
        print(str(dns)," DNS in CMDB +")  
        
def vm_ip_search (ip):
    API = '0a66e9aecdc8c49191309691d8c340b920a060b3'
    ip = ip
    endpoint = 'https://cmdb.netledger.com/api/ipaddresses/?address={ip}'
    url = endpoint.format(ip=ip)
    headers = {'Authorization': 'Token '+ API}
    response = requests.get(url, headers=headers)
    result = response.json()
    count = result['count']
    if count == 1:
        print(str(ip) + " IP IN CMDB ")
        workbook.vm['IP in CMDB'][workbook.vm['IP'] == ip] = 'IP in CMDB'
    elif count < 1:
        print(str(ip) + " IP NOT in CMDB ")

        workbook.vm['IP in CMDB'][workbook.vm['IP'] == ip] = 'IP NOT in CMDB'
    elif count > 1:
        print(str(ip) + " IP IN CMDB + ")

def vm_dns_search (dns):
    API = '0a66e9aecdc8c49191309691d8c340b920a060b3'
    dns = dns
    endpoint = 'https://cmdb.netledger.com/api/dc-hosts/?hostname={dns}'
    url = endpoint.format(dns=dns)
    headers = {'Authorization': 'Token '+ API}
    response = requests.get(url, headers=headers)
    result = response.json()
    count = result['count']
    if count == 1:
        workbook.vm['DNS in CMDB'][workbook.vm['DNS'] == dns] = 'DNS in CMDB'
        print(str(dns)," DNS in cmdb")
    elif count < 1:
        workbook.vm['DNS in CMDB'][workbook.vm['DNS'] == dns] = 'DNS NOT in CMDB'
        print(str(dns)," DNS NOT in cmdb")
    elif count > 1:
        workbook.vm['DNS in CMDB'][workbook.vm['DNS'] == dns] = 'DNS in CMDB + '
        print(str(dns)," DNS in CMDB +")  
        
def vmhost_search (dns):
    API = '0a66e9aecdc8c49191309691d8c340b920a060b3'
    dns = dns
    endpoint = 'https://cmdb.netledger.com/api/virtual-servers/?hostname={dns}'
    url = endpoint.format(dns=dns)
    headers = {'Authorization': 'Token '+ API}
    response = requests.get(url, headers=headers)
    result = response.json()
    count = result['count']
    if count != 0:
        url = result['results'][0]['hypervisor']['hostname']
        workbook.vm['VM host URL'][workbook.vm['DNS'] == dns] = url
        print(url)
    elif count == 0 :
        workbook.vm['VM host URL'][workbook.vm['DNS'] == dns] = 'None'
        print('None')
    else:
        url = dns
        workbook.vm['VM host URL'][workbook.vm['DNS'] == dns] = url
        print(url)
    

def available_ip_search (ip):
    API = '0a66e9aecdc8c49191309691d8c340b920a060b3'
    ip = ip
    endpoint = 'https://cmdb.netledger.com/api/ipaddresses/?address={ip}'
    url = endpoint.format(ip=ip)
    headers = {'Authorization': 'Token '+ API}
    response = requests.get(url, headers=headers)
    result = response.json()
    count = result['count']
    if count == 1:
        print(str(ip) + " IP IN CMDB ")
        workbook.available['IP in CMDB'][workbook.available['IP'] == ip] = 'IP in CMDB'
    elif count < 1:
        print(str(ip) + " IP NOT in CMDB ")

        workbook.available['IP in CMDB'][workbook.available['IP'] == ip] = 'IP NOT in CMDB'
    elif count > 1:
        print(str(ip) + " IP IN CMDB + ")

def available_dns_search (dns):
    API = '0a66e9aecdc8c49191309691d8c340b920a060b3'
    dns = dns
    endpoint = 'https://cmdb.netledger.com/api/dc-hosts/?hostname={dns}'
    url = endpoint.format(dns=dns)
    headers = {'Authorization': 'Token '+ API}
    response = requests.get(url, headers=headers)
    result = response.json()
    count = result['count']
    if count == 1:
        workbook.available['DNS in CMDB'][workbook.available['DNS'] == dns] = 'DNS in CMDB'
        print(str(dns)," DNS in cmdb")
    elif count < 1:
        workbook.available['DNS in CMDB'][workbook.available['DNS'] == dns] = 'DNS NOT in CMDB'
        print(str(dns)," DNS NOT in cmdb")
    elif count > 1:
        workbook.available['DNS in CMDB'][workbook.available['DNS'] == dns] = 'DNS in CMDB + '
        print(str(dns)," DNS in CMDB +")  
start = time.perf_counter()
print("Checking VPN  Connection...")

if connection_check().status_code == 200:
    print("VPN Connection GOOD")
else:
    while connection_check().status_code != 200:
        checking = input("Please check VPN Connection and then press 'Enter' to continue\n")
        connection_check().status_code == 200
        
# file_name = input("Enter Filename (don't include the file extension): \n")
file_name = "CMDB_queried_data_test"
print("Loading Workbook")
workbook = pd.ExcelFile(file_name+'.xlsx')
#Start time
start = time.perf_counter()
print("Loading Main tab")
workbook.main = pd.read_excel(workbook,sheet_name='Main')
#Main Tab IP CMDB Query
workbook.main['IP in CMDB']=''
#Main Tab DNS CMDB Query   
workbook.main['DNS in CMDB']=''
workbook.duplicate_IP = pd.read_excel(workbook,sheet_name='Duplicate IP')
print("Loading Duplicate IP tab")
workbook.duplicate_DNS = pd.read_excel(workbook,sheet_name='Duplicate DNS')
print("Loading Duplicate DNS tab")
workbook.bronto = pd.read_excel(workbook,sheet_name='Bronto')
print("Loading Bronto tab")
workbook.vpn = pd.read_excel(workbook,sheet_name='VPN')
print("Loading VPN tab")
print('VPN CMDB SEARCH')
#VPN Tab IP CMDB Query
workbook.vpn['IP in CMDB']=''
#VPN Tab DNS CMDB Query   
workbook.vpn['DNS in CMDB']=''
workbook.none = pd.read_excel(workbook,sheet_name='None')
print("Loading None tab")
print('None CMDB SEARCH')
#None Tab IP CMDB Query
workbook.none['IP in CMDB']=''
#None Tab DNS CMDB Query   
workbook.none['DNS in CMDB']=''
workbook.console = pd.read_excel(workbook,sheet_name='Console')
print("Loading Console tab")
print('Console CMDB SEARCH')
#console Tab IP CMDB Query
workbook.console['IP in CMDB']=''
#console Tab DNS CMDB Query   
workbook.console['DNS in CMDB']=''
workbook.interface = pd.read_excel(workbook,sheet_name='Interface')
print("Loading Interface tab")
print('Interface CMDB SEARCH')
#interface Tab IP CMDB Query
workbook.interface['IP in CMDB']=''
#interface Tab DNS CMDB Query   
workbook.interface['DNS in CMDB']=''
workbook.nouse = pd.read_excel(workbook,sheet_name='NoUse')
print("Loading NoUse tab")
print('NoUse CMDB SEARCH')
#nouse Tab IP CMDB Query
workbook.nouse['IP in CMDB']=''
#nouse Tab DNS CMDB Query   
workbook.nouse['DNS in CMDB']=''
workbook.vm = pd.read_excel(workbook,sheet_name='VM')
print("Loading VM tab")
print('VM CMDB SEARCH')
#VM Tab IP CMDB Query   
workbook.vm['IP in CMDB']=''
workbook.vm['DNS in CMDB']=''
workbook.vm['VM host URL']=''
workbook.available = pd.read_excel(workbook,sheet_name='Available')
print("Loading Available tab") 
print('available CMDB SEARCH')
#available Tab IP CMDB Query
workbook.available['IP in CMDB']=''
#available Tab DNS CMDB Query   
workbook.available['DNS in CMDB']=''

with concurrent.futures.ThreadPoolExecutor() as executor:
    executor.map(main_ip_search, workbook.main['IP'])
    executor.map(main_dns_search, workbook.main['DNS'])       
    executor.map(vpn_ip_search, workbook.vpn['IP'])
    executor.map(vpn_dns_search, workbook.vpn['DNS'])          
    executor.map(none_ip_search, workbook.none['IP'])
    executor.map(none_dns_search, workbook.none['DNS'])  
    executor.map(console_ip_search, workbook.console['IP'])
    executor.map(console_dns_search, workbook.console['DNS'])  
    executor.map(interface_ip_search, workbook.interface['IP'])
    executor.map(interface_dns_search, workbook.interface['DNS'])  
    executor.map(nouse_ip_search, workbook.nouse['IP'])
    executor.map(nouse_dns_search, workbook.nouse['DNS'])  
    executor.map(vm_ip_search, workbook.vm['IP'])
    executor.map(vm_dns_search, workbook.vm['DNS'])    
    executor.map(vmhost_search, workbook.vm['DNS'])        
    executor.map(available_ip_search, workbook.available['IP'])
    executor.map(available_dns_search, workbook.available['DNS'])  


# export_name = input('Enter File name (do not inclide file extension): \n')
export_name = "CMDB_queried_data_test"
with pd.ExcelWriter(export_name+'.xlsx') as writer:
    workbook.main.to_excel(writer,sheet_name='Main', index=False)
    workbook.duplicate_DNS.to_excel(writer,sheet_name='Duplicate DNS', index=False)
    workbook.duplicate_IP.to_excel(writer,sheet_name='Duplicate IP', index=False)
    workbook.bronto.to_excel(writer,sheet_name='Bronto', index=False)
    workbook.vpn.to_excel(writer,sheet_name='VPN', index=False)
    workbook.none.to_excel(writer,sheet_name='None', index=False)
    workbook.console.to_excel(writer,sheet_name='Console', index=False)
    workbook.interface.to_excel(writer,sheet_name='Interface', index=False)
    workbook.nouse.to_excel(writer,sheet_name='NoUse', index=False)
    workbook.vm.to_excel(writer,sheet_name='VM', index=False)
    workbook.available.to_excel(writer,sheet_name='Available', index=False)
    
print(export_name +" has been exported")
stop = time.perf_counter()

print("It took "+str(round(stop-start,2))+" seconds for the script to complete")