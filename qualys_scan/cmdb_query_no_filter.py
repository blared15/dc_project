import pandas as pd
import shutil,time, os.path,requests
import ipaddress
import concurrent.futures
pd.options.mode.chained_assignment = None
columns = shutil.get_terminal_size().columns

def connection_check():
    API = '0a66e9aecdc8c49191309691d8c340b920a060b3'
    endpoint = 'https://cmdb.netledger.com/api/dc-hosts/'
    headers = {'Authorization': 'Token '+ API}
    response = requests.get(endpoint, headers=headers)
    return response

cloud_file = input("enter CLOUD file name location ex: ~/Downloads/qualys-cloud-hostlist (no file extension): \n")
cloud_check = os.path.isfile(cloud_file+'.csv')
while cloud_check == False:
    cloud_file = input("enter cloud file name location ex: ~/Downloads/qualys-cloud-hostlist (no file extension): \n")
    cloud_check = os.path.isfile(cloud_file+'.csv')
print("Cloud file found: " + str(cloud_check))

pcp_file = input("enter PCP file name location ex: ~/Downloads/qualys-pcp-hostlist (no file extension): \n")
pcp_check = os.path.isfile(pcp_file+'.csv')
while pcp_check == False:
    pcp_file = input("enter PCP file name location ex: ~/Downloads/qualys-pcp-hostlist (no file extension): \n")
    pcp_check = os.path.isfile(pcp_file+'.csv')
print("PCP file found: " + str(pcp_check))

start = time.perf_counter()
#converting file to pd
cloud = pd.read_csv(cloud_file+'.csv')
#add from to determine where the data came from
cloud['From']=''
for x in cloud["From"]:
    cloud["From"] = "Cloud"
    
pcp = pd.read_csv(pcp_file+'.csv')
#add from to determine where the data came from
pcp['From']=''
for x in pcp["From"]:
    pcp["From"] = "PCP"
    
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
        merg['IP in CMDB'][merg['IP'] == ip] = 'IP in CMDB'
    elif count < 1:
        print(str(ip) + " IP NOT in CMDB ")

        merg['IP in CMDB'][merg['IP'] == ip] = 'IP NOT in CMDB'
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
        merg['DNS in CMDB'][merg['DNS'] == dns] = 'DNS in CMDB'
        print(str(dns)," DNS in cmdb")
    elif count < 1:
        merg['DNS in CMDB'][merg['DNS'] == dns] = 'DNS NOT in CMDB'
        print(str(dns)," DNS NOT in cmdb")
    elif count > 1:
        merg['DNS in CMDB'][merg['DNS'] == dns] = 'DNS in CMDB + '
        print(str(dns)," DNS in CMDB +")  
        
print("Checking VPN  Connection...")

if connection_check().status_code == 200:
    print("VPN Connection GOOD")
else:
    while connection_check().status_code != 200:
        checking = input("Please check VPN Connection and then press 'Enter' to continue\n")
        connection_check().status_code == 200
        
print("Merging data")
merg = cloud.append(pcp)
print("Data count after merge: " + str(len(merg)))


merg['IP in CMDB']=''
merg['DNS in CMDB']=''

with concurrent.futures.ThreadPoolExecutor() as executor:
    executor.map(main_ip_search, merg['IP'])
    executor.map(main_dns_search, merg['DNS']) 


with pd.ExcelWriter('CMDB_queried_data_no_filter.xlsx') as writer:
    merg.to_excel(writer,index=False)
    
stop = time.perf_counter()
print("It took "+str(round((stop-start)/60,2))+" minutes for the script to complete")