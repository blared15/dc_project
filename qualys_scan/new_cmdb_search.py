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

def bronto_ip_search (ip):
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
        workbook.bronto['IP in CMDB'][workbook.bronto['IP'] == ip] = 'IP in CMDB'
    elif count < 1:
        print(str(ip) + " IP NOT in CMDB ")

        workbook.bronto['IP in CMDB'][workbook.bronto['IP'] == ip] = 'IP NOT in CMDB'
    elif count > 1:
        print(str(ip) + " IP IN CMDB + ")

def bronto_dns_search (dns):
    API = '0a66e9aecdc8c49191309691d8c340b920a060b3'
    dns = dns
    endpoint = 'https://cmdb.netledger.com/api/dc-hosts/?hostname={dns}'
    url = endpoint.format(dns=dns)
    headers = {'Authorization': 'Token '+ API}
    response = requests.get(url, headers=headers)
    result = response.json()
    count = result['count']
    if count == 1:
        workbook.bronto['DNS in CMDB'][workbook.bronto['DNS'] == dns] = 'DNS in CMDB'
        print(str(dns)," DNS in cmdb")
    elif count < 1:
        workbook.bronto['DNS in CMDB'][workbook.bronto['DNS'] == dns] = 'DNS NOT in CMDB'
        print(str(dns)," DNS NOT in cmdb")
    elif count > 1:
        workbook.bronto['DNS in CMDB'][workbook.bronto['DNS'] == dns] = 'DNS in CMDB + '
        print(str(dns)," DNS in CMDB +")  

def openair_ip_search (ip):
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
        workbook.openair['IP in CMDB'][workbook.openair['IP'] == ip] = 'IP in CMDB'
    elif count < 1:
        print(str(ip) + " IP NOT in CMDB ")

        workbook.openair['IP in CMDB'][workbook.openair['IP'] == ip] = 'IP NOT in CMDB'
    elif count > 1:
        print(str(ip) + " IP IN CMDB + ")

def openair_dns_search (dns):
    API = '0a66e9aecdc8c49191309691d8c340b920a060b3'
    dns = dns
    endpoint = 'https://cmdb.netledger.com/api/dc-hosts/?hostname={dns}'
    url = endpoint.format(dns=dns)
    headers = {'Authorization': 'Token '+ API}
    response = requests.get(url, headers=headers)
    result = response.json()
    count = result['count']
    if count == 1:
        workbook.openair['DNS in CMDB'][workbook.openair['DNS'] == dns] = 'DNS in CMDB'
        print(str(dns)," DNS in cmdb")
    elif count < 1:
        workbook.openair['DNS in CMDB'][workbook.openair['DNS'] == dns] = 'DNS NOT in CMDB'
        print(str(dns)," DNS NOT in cmdb")
    elif count > 1:
        workbook.openair['DNS in CMDB'][workbook.openair['DNS'] == dns] = 'DNS in CMDB + '
        print(str(dns)," DNS in CMDB +")  
        
def oci_ip_search (ip):
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
        workbook.oci['IP in CMDB'][workbook.oci['IP'] == ip] = 'IP in CMDB'
    elif count < 1:
        print(str(ip) + " IP NOT in CMDB ")

        workbook.oci['IP in CMDB'][workbook.oci['IP'] == ip] = 'IP NOT in CMDB'
    elif count > 1:
        print(str(ip) + " IP IN CMDB + ")

def oci_dns_search (dns):
    API = '0a66e9aecdc8c49191309691d8c340b920a060b3'
    dns = dns
    endpoint = 'https://cmdb.netledger.com/api/dc-hosts/?hostname={dns}'
    url = endpoint.format(dns=dns)
    headers = {'Authorization': 'Token '+ API}
    response = requests.get(url, headers=headers)
    result = response.json()
    count = result['count']
    if count == 1:
        workbook.oci['DNS in CMDB'][workbook.oci['DNS'] == dns] = 'DNS in CMDB'
        print(str(dns)," DNS in cmdb")
    elif count < 1:
        workbook.oci['DNS in CMDB'][workbook.oci['DNS'] == dns] = 'DNS NOT in CMDB'
        print(str(dns)," DNS NOT in cmdb")
    elif count > 1:
        workbook.oci['DNS in CMDB'][workbook.oci['DNS'] == dns] = 'DNS in CMDB + '
        print(str(dns)," DNS in CMDB +")  

def payroll_ip_search (ip):
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
        workbook.payroll['IP in CMDB'][workbook.payroll['IP'] == ip] = 'IP in CMDB'
    elif count < 1:
        print(str(ip) + " IP NOT in CMDB ")

        workbook.payroll['IP in CMDB'][workbook.payroll['IP'] == ip] = 'IP NOT in CMDB'
    elif count > 1:
        print(str(ip) + " IP IN CMDB + ")

def payroll_dns_search (dns):
    API = '0a66e9aecdc8c49191309691d8c340b920a060b3'
    dns = dns
    endpoint = 'https://cmdb.netledger.com/api/dc-hosts/?hostname={dns}'
    url = endpoint.format(dns=dns)
    headers = {'Authorization': 'Token '+ API}
    response = requests.get(url, headers=headers)
    result = response.json()
    count = result['count']
    if count == 1:
        workbook.payroll['DNS in CMDB'][workbook.payroll['DNS'] == dns] = 'DNS in CMDB'
        print(str(dns)," DNS in cmdb")
    elif count < 1:
        workbook.payroll['DNS in CMDB'][workbook.payroll['DNS'] == dns] = 'DNS NOT in CMDB'
        print(str(dns)," DNS NOT in cmdb")
    elif count > 1:
        workbook.payroll['DNS in CMDB'][workbook.payroll['DNS'] == dns] = 'DNS in CMDB + '
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
        
def blank_ip_search (ip):
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
        workbook.blank['IP in CMDB'][workbook.blank['IP'] == ip] = 'IP in CMDB'
    elif count < 1:
        print(str(ip) + " IP NOT in CMDB ")

        workbook.blank['IP in CMDB'][workbook.blank['IP'] == ip] = 'IP NOT in CMDB'
    elif count > 1:
        print(str(ip) + " IP IN CMDB + ")

def blank_dns_search (dns):
    API = '0a66e9aecdc8c49191309691d8c340b920a060b3'
    dns = dns
    endpoint = 'https://cmdb.netledger.com/api/dc-hosts/?hostname={dns}'
    url = endpoint.format(dns=dns)
    headers = {'Authorization': 'Token '+ API}
    response = requests.get(url, headers=headers)
    result = response.json()
    count = result['count']
    if count == 1:
        workbook.blank['DNS in CMDB'][workbook.blank['DNS'] == dns] = 'DNS in CMDB'
        print(str(dns)," DNS in cmdb")
    elif count < 1:
        workbook.blank['DNS in CMDB'][workbook.blank['DNS'] == dns] = 'DNS NOT in CMDB'
        print(str(dns)," DNS NOT in cmdb")
    elif count > 1:
        workbook.blank['DNS in CMDB'][workbook.blank['DNS'] == dns] = 'DNS in CMDB + '
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
        workbook.vm['VM host URL'][workbook.vm['DNS'] == dns] = 'blank'
        print('blank')
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
        
file_name = input("Enter merged filename ex: New_Data_Latest (don't include the file extension): \n")
# file_name = "CMDB_queried_data_test"
#Start time
start = time.perf_counter()
print("Loading Workbook")

workbook = pd.ExcelFile(file_name+'.xlsx')

print("Loading Main tab")
workbook.main = pd.read_excel(workbook,sheet_name='Main')
#Main Tab IP CMDB Query
workbook.main['IP in CMDB']=''
#Main Tab DNS CMDB Query   
workbook.main['DNS in CMDB']=''

print("Loading Duplicate IP tab")
workbook.duplicate_IP = pd.read_excel(workbook,sheet_name='Duplicate IP')

print("Loading Duplicate DNS tab")
workbook.duplicate_DNS = pd.read_excel(workbook,sheet_name='Duplicate DNS')

print("Loading Bronto tab")
workbook.bronto = pd.read_excel(workbook,sheet_name='Bronto')
#Bronto Tab IP CMDB Query
workbook.bronto['IP in CMDB']=''
#Bronto Tab DNS CMDB Query   
workbook.bronto['DNS in CMDB']=''

print("Loading OpenAir tab")
workbook.openair = pd.read_excel(workbook,sheet_name='OpenAir')
#OpenAir Tab IP CMDB Query
workbook.openair['IP in CMDB']=''
#OpenAir Tab DNS CMDB Query   
workbook.openair['DNS in CMDB']=''

print("Loading OCI tab")
workbook.oci = pd.read_excel(workbook,sheet_name='OCI')
#OCI Tab IP CMDB Query
workbook.oci['IP in CMDB']=''
#OCI Tab DNS CMDB Query   
workbook.oci['DNS in CMDB']=''

print("Loading Payroll tab")
workbook.payroll = pd.read_excel(workbook,sheet_name='OCI')
#Payroll Tab IP CMDB Query
workbook.payroll['IP in CMDB']=''
#Payroll Tab DNS CMDB Query   
workbook.payroll['DNS in CMDB']=''

workbook.vpn = pd.read_excel(workbook,sheet_name='VPN')
print("Loading VPN tab")
print('VPN CMDB SEARCH')
#VPN Tab IP CMDB Query
workbook.vpn['IP in CMDB']=''
#VPN Tab DNS CMDB Query   
workbook.vpn['DNS in CMDB']=''

print("Loading None tab")
workbook.blank = pd.read_excel(workbook,sheet_name='None')
print('None CMDB SEARCH')
#None Tab IP CMDB Query
workbook.blank['IP in CMDB']=''
#blank Tab DNS CMDB Query   
workbook.blank['DNS in CMDB']=''

print("Loading Console tab")
workbook.console = pd.read_excel(workbook,sheet_name='Console')
print('Console CMDB SEARCH')
#console Tab IP CMDB Query
workbook.console['IP in CMDB']=''
#console Tab DNS CMDB Query   
workbook.console['DNS in CMDB']=''

print("Loading Interface tab")
workbook.interface = pd.read_excel(workbook,sheet_name='Interface')
print('Interface CMDB SEARCH')
#interface Tab IP CMDB Query
workbook.interface['IP in CMDB']=''
#interface Tab DNS CMDB Query   
workbook.interface['DNS in CMDB']=''

print("Loading NoUse tab")
workbook.nouse = pd.read_excel(workbook,sheet_name='NoUse')
print('NoUse CMDB SEARCH')
#nouse Tab IP CMDB Query
workbook.nouse['IP in CMDB']=''
#nouse Tab DNS CMDB Query   
workbook.nouse['DNS in CMDB']=''

print("Loading VM tab")
workbook.vm = pd.read_excel(workbook,sheet_name='VM')
print('VM CMDB SEARCH')
#VM Tab IP CMDB Query   
workbook.vm['IP in CMDB']=''
workbook.vm['DNS in CMDB']=''
workbook.vm['VM host URL']=''

print("Loading Available tab") 
workbook.available = pd.read_excel(workbook,sheet_name='Available')
print('available CMDB SEARCH')
#available Tab IP CMDB Query
workbook.available['IP in CMDB']=''
#available Tab DNS CMDB Query   
workbook.available['DNS in CMDB']=''




with concurrent.futures.ThreadPoolExecutor() as executor:
    executor.map(main_ip_search, workbook.main['IP'])
    executor.map(main_dns_search, workbook.main['DNS']) 
    
    executor.map(bronto_ip_search, workbook.bronto['IP'])
    executor.map(bronto_dns_search, workbook.bronto['DNS'])  
    
    executor.map(openair_ip_search, workbook.openair['IP'])
    executor.map(openair_dns_search, workbook.openair['DNS'])  
    
    executor.map(oci_ip_search, workbook.oci['IP'])
    executor.map(oci_dns_search, workbook.oci['DNS']) 
    
    executor.map(payroll_ip_search, workbook.payroll['IP'])
    executor.map(payroll_dns_search, workbook.payroll['DNS']) 
    
    executor.map(vpn_ip_search, workbook.vpn['IP'])
    executor.map(vpn_dns_search, workbook.vpn['DNS'])   

    executor.map(blank_ip_search, workbook.blank['IP'])
    executor.map(blank_dns_search, workbook.blank['DNS'])  
    
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
# export_name = "CMDB_queried_data_test"
with pd.ExcelWriter('CMDB_queried_data_latest.xlsx') as writer:
    workbook.main.to_excel(writer,sheet_name='Main', index=False)
    workbook.duplicate_DNS.to_excel(writer,sheet_name='Duplicate DNS', index=False)
    workbook.duplicate_IP.to_excel(writer,sheet_name='Duplicate IP', index=False)
    workbook.bronto.to_excel(writer,sheet_name='Bronto', index=False)
    workbook.oci.to_excel(writer,sheet_name='OCI', index=False)
    workbook.openair.to_excel(writer,sheet_name='OpenAir', index=False)
    workbook.payroll.to_excel(writer,sheet_name='Payroll', index=False)
    workbook.vpn.to_excel(writer,sheet_name='VPN', index=False)
    workbook.blank.to_excel(writer,sheet_name='None', index=False)
    workbook.console.to_excel(writer,sheet_name='Console', index=False)
    workbook.interface.to_excel(writer,sheet_name='Interface', index=False)
    workbook.nouse.to_excel(writer,sheet_name='NoUse', index=False)
    workbook.vm.to_excel(writer,sheet_name='VM', index=False)
    workbook.available.to_excel(writer,sheet_name='Available', index=False)


print("CMDB_queried_data_latest has been exported")
stop = time.perf_counter()

print("It took "+str(round((stop-start)/60,2))+" minutes for the script to complete")