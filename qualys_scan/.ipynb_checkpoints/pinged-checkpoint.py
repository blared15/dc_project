# To add a new cell, type '# %%'
# To add a new markdown cell, type '# %% [markdown]'
# %%
import os, time, concurrent.futures, shutil,requests, subprocess
import pandas as pd
columns = shutil.get_terminal_size().columns


# %%
title = (r"""
  _____  _____     ____                            
 |_   _||  __ \   / __ \                           
   | |  | |__) | | |  | | _   _   ___  _ __  _   _ 
   | |  |  ___/  | |  | || | | | / _ \| '__|| | | |
  _| |_ | |      | |__| || |_| ||  __/| |   | |_| |
 |_____||_|       \___\_\ \__,_| \___||_|    \__, |
                                              __/ |
                                             |___/ 
""")

# %%
print(title)
print("*".center(columns,'*'))       
print ("This program check if the ip address response to ping.".center(columns))
print("Make sure that you are connected to the VPN.".center(columns))
print("This program will take about ~60 minutes to finish ping the IP list.".center(columns))
print("*".center(columns,'*'))      


# %%
def main_ping(ip):
    response = subprocess.call(['ping','-c 2','-t 2',f'{ip}']) == 0
    if response == True:
        print(f"UP {ip} Ping Successful")
        data['Ping'][data['IP']==ip]='UP'
    else:
        print(f"DOWN {ip} Ping Unsuccessful")
        data['Ping'][data['IP']==ip]='DOWN'
def main_lookup(ip):
    response = os.popen("nslookup "+ ip+" | awk 'NR==5{print $4}'").read()
    if response == '=\n':
        response = os.popen("nslookup "+ ip+" | awk 'NR==6{print $4}'").read()
        response = response.rstrip()
        print(response)
        data['CNAME'][data['IP']==ip]=response
    elif response == '\n':
        print(f"{ip} No DNS")
        data['CNAME'][data['IP']==ip]='None'
    else:
        print(f"{ip} {response}")
        response = response.rstrip()
        print(response)
        data['CNAME'][data['IP']==ip]=response

def file_check(file_name):
    check = os.path.isfile(file_name)
    file_name = file_name +".xlsx"
    while check == False:
        file_name = input("File name not found.\n Enter XLSX file that contains the merged files (no tabs) without the file extension. ex: ~/Downloads/raw_data :\n")
        check = os.path.isfile(file_name)
    print(" file found: " + str(check))    

def connection_check():
    API = '0a66e9aecdc8c49191309691d8c340b920a060b3'
    endpoint = 'https://cmdb.netledger.com/api/dc-hosts/'
    headers = {'Authorization': 'Token '+ API}
    response = requests.get(endpoint, headers=headers)
    print (response)


# %%
file_name = input("Enter XLSX file that contains the merged files (no tabs) without the file extension. ex: ~/Downloads/raw_data :\n")
file_name = file_name +".xlsx"
file_check(file_name)
print("Checking VPN connection: ")
connection_check()

start = time.perf_counter()
data = pd.read_excel(file_name)
data['Ping']=''
data['CNAME']=''


# %%
with concurrent.futures.ThreadPoolExecutor() as executor:
    executor.map(main_ping,data['IP'])
    executor.map(main_lookup,data['IP'])
    
y = data[data['Ping'].isnull()].index

while  len(y) !=0:
    for index in y:
        main_ping(data['IP'][data['IP'].index==index].to_string(index=False).lstrip())
    y = data[data['Ping'].isnull()].index


# %%
with pd.ExcelWriter('latest_sample_pinged_data.xlsx') as writer:
    data.to_excel(writer, index=False)

stop = time.perf_counter()

print("It took "+str(round((stop-start)/60,2))+" minutes for the script to complete")


