# To add a new cell, type '# %%'
# To add a new markdown cell, type '# %% [markdown]'
# %%
import time, os, requests,shutil
import pandas as pd
import concurrent.futures
from datetime import datetime, timedelta
import datetime
columns = shutil.get_terminal_size().columns


# %%
def file_check(file_name, types):
    check = os.path.isfile(file_name)
    file_name = file_name +".csv"
    while check == False:
        file_name = input("File name not found.\n Enter "+types+" filename including location without the file extension. ex: ~/Downloads/qualys-"+types+"-hostlist:\n")
        check = os.path.isfile(file_name)
    print(types + " file found: " + str(check))    

def connection_check():
    API = '0a66e9aecdc8c49191309691d8c340b920a060b3'
    endpoint = 'https://cmdb.netledger.com/api/dc-hosts/'
    headers = {'Authorization': 'Token '+ API}
    response = requests.get(endpoint, headers=headers)
    return response

def ip_search (ip):
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
        merg['IP in CMDB'][merg['IP'] == ip] = 'IP in CMDB'
    elif count < 1:
        print(str(ip) + " IP NOT in CMDB ")
        merg['IP in CMDB'][merg['IP'] == ip] = 'IP NOT in CMDB'
    elif count > 1:
        print(str(ip) + " IP IN CMDB + ")
        merg['IP in CMDB'][merg['IP'] == ip] = 'IP in CMDB+'

def outdated(dns):
    API = '0a66e9aecdc8c49191309691d8c340b920a060b3'
    url = f'https://cmdb.netledger.com/api/dc-hosts/?hostname=Outdated%20({dns})'
    headers = {'Authorization': 'Token '+ API}
    response = requests.get(url, headers=headers)
    result = response.json()
    count = result['count']
    if count == 1:
            merg['DNS in CMDB'][merg['DNS'] == dns] = 'DNS in CMDB - Outdated'
            print(str(dns),'DNS in CMDB - Outdated')
    elif count < 1:
            merg['DNS in CMDB'][merg['DNS'] == dns] = 'DNS NOT in CMDB'
            print(str(dns)," DNS NOT in cmdb")  
    

def dns_search (dns):
    API = '0a66e9aecdc8c49191309691d8c340b920a060b3'
    url = f'https://cmdb.netledger.com/api/dc-hosts/?hostname={dns}'
    headers = {'Authorization': 'Token '+ API}
    response = requests.get(url, headers=headers)
    result = response.json()
    count = result['count']
    if count == 0:
        count = outdated(dns)
    
    elif count == 1:
        merg['DNS in CMDB'][merg['DNS'] == dns] = 'DNS in CMDB'
        print(str(dns)," DNS in cmdb")
    elif count > 1:
        merg['DNS in CMDB'][merg['DNS'] == dns] = 'DNS in CMDB + '
        
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
        merg['VM host URL'][merg['DNS'] == dns] = url
        print(url)
    elif count == 0 :
        merg['VM host URL'][merg['DNS'] == dns] = 'blank'
        print('blank')
    else:
        url = dns
        merg['VM host URL'][merg['DNS'] == dns] = url
        print(url) 
        
def vm_check(dns):
    if '.f.' in dns:
        vmhost_search(dns)
    elif '.m.' in dns:
        vmhost_search(dns)
    elif '.snap.' in dns:
        vmhost_search(dns)
    else:
        pass



# %%
title=(r"""
                         __  __                                              ____                                
                        |  \/  |                                  ___       / __ \                               
                        | \  / |   ___   _ __    __ _    ___     ( _ )     | |  | |  _   _    ___   _ __   _   _ 
                        | |\/| |  / _ \ | '__|  / _` |  / _ \    / _ \/\   | |  | | | | | |  / _ \ | '__| | | | |
                        | |  | | |  __/ | |    | (_| | |  __/   | (_>  <   | |__| | | |_| | |  __/ | |    | |_| |
                        |_|  |_|  \___| |_|     \__, |  \___|    \___/\/    \___\_\  \__,_|  \___| |_|     \__, |
                                                __/ |                                                      __/ |
                                                |___/                                                      |___/ 

                     """)


# %%
print(title)
print("*".center(columns,'*'))       
print ("This program will merge 2 Qualys Scan CSV file and query CMDB.".center(columns))
print("Make sure that you are connected to the VPN.".center(columns))
print("This program will take about ~30 minutes to finish querying CMDB.".center(columns))
print("*".center(columns,'*'))       


# %%
if connection_check().status_code == 200:
    print("VPN Connection GOOD")
else:
    while connection_check().status_code != 200:
        checking = input("Please check VPN Connection and then press 'Enter' to continue\n")
        connection_check().status_code == 200

    


# %%
#select pcp file
pcp_file = input('Enter pcp file name location. Do not include the file extension. ex: ~/Downloads/qualys-pcp-hostlist: ')
pcp_file = pcp_file + ".csv"
file_check(pcp_file,'pcp')


# %%
#select cloud file
cloud_file = input('Enter cloud file name location. Do not include the file extension. ex: ~/Downloads/qualys-cloud-hostlist: ')
cloud_file = cloud_file + ".csv"
file_check(cloud_file,'cloud')


# %%
start = time.perf_counter()
#converting file to pd
print('Loading Cloud file')
cloud = pd.read_csv(cloud_file)
cloud['From']=''
for x in cloud["From"]:
    cloud["From"] = "Cloud"
print('Loading PCP file')
pcp = pd.read_csv(pcp_file)
pcp['From']=''
for x in pcp["From"]:
    pcp["From"] = "PCP"


# %%
print("Merging data")
merg = cloud.append(pcp)
print("Data count after merge: " + str(len(merg)))
init_count = len(merg)
time.sleep(2)

#change last scanned date to_date
merg['Last Scanned Date']=pd.to_datetime(merg['Last Scanned Date'], format= '%Y-%m-%d').dt.date
# %%
#Adding columns
merg["IP in CMDB"]=""
merg['DNS in CMDB']=''
merg['VM host URL']=''



# %%



# %%
#query starts
with concurrent.futures.ThreadPoolExecutor() as executor:
    executor.map(ip_search, merg['IP'])
    executor.map(dns_search, merg['DNS'])
    executor.map(vm_check, merg['DNS'])

x = merg[merg['IP in CMDB'].isnull()].index
y = merg[merg['DNS in CMDB'].isnull()].index
while len(x) != 0 or len(y) !=0:
    for index in x:
        ip_search(merg['IP'][merg['IP'].index==index].to_string(index=False).lstrip())
        for index in y:
            dns_search(merg['DNS'][merg['DNS'].index==index].to_string(index=False).lstrip())
    x = merg[merg['IP in CMDB'].isnull()].index
    y = merg[merg['DNS in CMDB'].isnull()].index        

#older than 30days from last scanned date
old = merg[merg['Last Scanned Date'] < datetime.date(2021,1,19)- pd.to_timedelta("30day") ].sort_values('Last Scanned Date' ,ascending=False)

#30 days from last scanned date
merg = merg[merg['Last Scanned Date'] > datetime.date(2021,1,19)- pd.to_timedelta("30day") ].sort_values('Last Scanned Date' ,ascending=False)

# %%
with pd.ExcelWriter('CMDB_Queried_Data_Latest.xlsx') as writer:
    merg.to_excel(writer, index=False)
    
with pd.ExcelWriter('CMDB_Queried_Data_Latest_30_plus.xlsx') as writer:
    old.to_excel(writer, index=False)

print("CMDB_Queried_Data_Latest_Test.xlsx and CMDB_Queried_Data_Latest_30_plus.xlsx has been exported")
stop = time.perf_counter()

print("It took "+str(round((stop-start)/60,2))+" minutes for the script to complete")


# %%



