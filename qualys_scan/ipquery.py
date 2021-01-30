import pandas as pd
import os, time
import concurrent.futures
# os.popen("nslookup 172.17.14.40 | awk 'NR==5{print $4}'").read()
def main_ping(ip):
    response = os.popen(f"ping -c 1 {ip}").read()
    if "Received = 4" in response:
        print(f"UP {ip} Ping Successful")
        workbook.main['Ping'][workbook.main['IP']==ip]='UP'
    else:
        print(f"DOWN {ip} Ping Unsuccessful")
        workbook.main['Ping'][workbook.main['IP']==ip]='DOWN'
def main_lookup(ip):
    response = os.popen("nslookup "+ ip+" | awk 'NR==5{print $4}'").read()
    if response == '\n':
        print(f"{ip} No DNS")
        workbook.main['CNAME'][workbook.main['IP']==ip]='None'
    else:
        print(f"{ip} {response}")
        response = response.rstrip()
        workbook.main['CNAME'][workbook.main['IP']==ip]=response

def ls_ping(ip):
    response = os.popen(f"ping -c 1 {ip}").read()
    if "Received = 4" in response:
        print(f"UP {ip} Ping Successful")
        workbook.ls['Ping'][workbook.ls['IP']==ip]='UP'
    else:
        print(f"DOWN {ip} Ping Unsuccessful")
        workbook.ls['Ping'][workbook.ls['IP']==ip]='DOWN'
def ls_lookup(ip):
    response = os.popen("nslookup "+ ip+" | awk 'NR==5{print $4}'").read()
    if response == '\n':
        print(f"{ip} No DNS")
        workbook.ls['CNAME'][workbook.ls['IP']==ip]='None'
    else:
        print(f"{ip} {response}")
        response = response.rstrip()
        workbook.ls['CNAME'][workbook.ls['IP']==ip]=response
        
def cy_ping(ip):
    response = os.popen(f"ping -c 1 {ip}").read()
    if "Received = 4" in response:
        print(f"UP {ip} Ping Successful")
        workbook.cy['Ping'][workbook.cy['IP']==ip]='UP'
    else:
        print(f"DOWN {ip} Ping Unsuccessful")
        workbook.cy['Ping'][workbook.cy['IP']==ip]='DOWN'
def cy_lookup(ip):
    response = os.popen("nslookup "+ ip+" | awk 'NR==5{print $4}'").read()
    if response == '\n':
        print(f"{ip} No DNS")
        workbook.cy['CNAME'][workbook.cy['IP']==ip]='None'
    else:
        print(f"{ip} {response}")
        response = response.rstrip()
        workbook.cy['CNAME'][workbook.cy['IP']==ip]=response        

        
def pdu_ping(ip):
    response = os.popen(f"ping -c 1 {ip}").read()
    if "Received = 4" in response:
        print(f"UP {ip} Ping Successful")
        workbook.pdu['Ping'][workbook.pdu['IP']==ip]='UP'
    else:
        print(f"DOWN {ip} Ping Unsuccessful")
        workbook.pdu['Ping'][workbook.pdu['IP']==ip]='DOWN'
def pdu_lookup(ip):
    response = os.popen("nslookup "+ ip+" | awk 'NR==5{print $4}'").read()
    if response == '\n':
        print(f"{ip} No DNS")
        workbook.pdu['CNAME'][workbook.pdu['IP']==ip]='None'
    else:
        print(f"{ip} {response}")
        response = response.rstrip()
        workbook.pdu['CNAME'][workbook.pdu['IP']==ip]=response

def pub_ping(ip):
    response = os.popen(f"ping -c 1 {ip}").read()
    if "Received = 4" in response:
        print(f"UP {ip} Ping Successful")
        workbook.pub['Ping'][workbook.pub['IP']==ip]='UP'
    else:
        print(f"DOWN {ip} Ping Unsuccessful")
        workbook.pub['Ping'][workbook.pub['IP']==ip]='DOWN'
def pub_lookup(ip):
    response = os.popen("nslookup "+ ip+" | awk 'NR==5{print $4}'").read()
    if response == '\n':
        print(f"{ip} No DNS")
        workbook.pub['CNAME'][workbook.pub['IP']==ip]='None'
    else:
        print(f"{ip} {response}")
        response = response.rstrip()
        workbook.pub['CNAME'][workbook.pub['IP']==ip]=response
        
def bo1_ping(ip):
    response = os.popen(f"ping -c 1 {ip}").read()
    if "Received = 4" in response:
        print(f"UP {ip} Ping Successful")
        workbook.bo1['Ping'][workbook.bo1['IP']==ip]='UP'
    else:
        print(f"DOWN {ip} Ping Unsuccessful")
        workbook.bo1['Ping'][workbook.bo1['IP']==ip]='DOWN'
def bo1_lookup(ip):
    response = os.popen("nslookup "+ ip+" | awk 'NR==5{print $4}'").read()
    if response == '\n':
        print(f"{ip} No DNS")
        workbook.bo1['CNAME'][workbook.bo1['IP']==ip]='None'
    else:
        print(f"{ip} {response}")
        response = response.rstrip()
        workbook.bo1['CNAME'][workbook.bo1['IP']==ip]=response

def agent_ping(ip):
    response = os.popen(f"ping -c 1 {ip}").read()
    if "Received = 4" in response:
        print(f"UP {ip} Ping Successful")
        workbook.agent['Ping'][workbook.agent['IP']==ip]='UP'
    else:
        print(f"DOWN {ip} Ping Unsuccessful")
        workbook.agent['Ping'][workbook.agent['IP']==ip]='DOWN'
def agent_lookup(ip):
    response = os.popen("nslookup "+ ip+" | awk 'NR==5{print $4}'").read()
    if response == '\n':
        print(f"{ip} No DNS")
        workbook.agent['CNAME'][workbook.agent['IP']==ip]='None'
    else:
        print(f"{ip} {response}")
        response = response.rstrip()
        workbook.agent['CNAME'][workbook.agent['IP']==ip]=response
        
def bronto_ping(ip):
    response = os.popen(f"ping -c 1 {ip}").read()
    if "Received = 4" in response:
        print(f"UP {ip} Ping Successful")
        workbook.bronto['Ping'][workbook.bronto['IP']==ip]='UP'
    else:
        print(f"DOWN {ip} Ping Unsuccessful")
        workbook.bronto['Ping'][workbook.bronto['IP']==ip]='DOWN'
def bronto_lookup(ip):
    response = os.popen("nslookup "+ ip+" | awk 'NR==5{print $4}'").read()
    if response == '\n':
        print(f"{ip} No DNS")
        workbook.bronto['CNAME'][workbook.bronto['IP']==ip]='None'
    else:
        print(f"{ip} {response}")
        response = response.rstrip()
        workbook.bronto['CNAME'][workbook.bronto['IP']==ip]=response

def openair_ping(ip):
    response = os.popen(f"ping -c 1 {ip}").read()
    if "Received = 4" in response:
        print(f"UP {ip} Ping Successful")
        workbook.openair['Ping'][workbook.openair['IP']==ip]='UP'
    else:
        print(f"DOWN {ip} Ping Unsuccessful")
        workbook.openair['Ping'][workbook.openair['IP']==ip]='DOWN'
def openair_lookup(ip):
    response = os.popen("nslookup "+ ip+" | awk 'NR==5{print $4}'").read()
    if response == '\n':
        print(f"{ip} No DNS")
        workbook.pdu['CNAME'][workbook.pdu['IP']==ip]='None'
    else:
        print(f"{ip} {response}")
        response = response.rstrip()
        workbook.pdu['CNAME'][workbook.pdu['IP']==ip]=response
        
def oci_ping(ip):
    response = os.popen(f"ping -c 1 {ip}").read()
    if "Received = 4" in response:
        print(f"UP {ip} Ping Successful")
        workbook.oci['Ping'][workbook.oci['IP']==ip]='UP'
    else:
        print(f"DOWN {ip} Ping Unsuccessful")
        workbook.oci['Ping'][workbook.oci['IP']==ip]='DOWN'
def oci_lookup(ip):
    response = os.popen("nslookup "+ ip+" | awk 'NR==5{print $4}'").read()
    if response == '\n':
        print(f"{ip} No DNS")
        workbook.oci['CNAME'][workbook.oci['IP']==ip]='None'
    else:
        print(f"{ip} {response}")
        response = response.rstrip()
        workbook.oci['CNAME'][workbook.oci['IP']==ip]=response
        
def payroll_ping(ip):
    response = os.popen(f"ping -c 1 {ip}").read()
    if "Received = 4" in response:
        print(f"UP {ip} Ping Successful")
        workbook.payroll['Ping'][workbook.payroll['IP']==ip]='UP'
    else:
        print(f"DOWN {ip} Ping Unsuccessful")
        workbook.payroll['Ping'][workbook.payroll['IP']==ip]='DOWN'
def payroll_lookup(ip):
    response = os.popen("nslookup "+ ip+" | awk 'NR==5{print $4}'").read()
    if response == '\n':
        print(f"{ip} No DNS")
        workbook.payroll['CNAME'][workbook.payroll['IP']==ip]='None'
    else:
        print(f"{ip} {response}")
        response = response.rstrip()
        workbook.payroll['CNAME'][workbook.payroll['IP']==ip]=response
        
def vpn_ping(ip):
    response = os.popen(f"ping -c 1 {ip}").read()
    if "Received = 4" in response:
        print(f"UP {ip} Ping Successful")
        workbook.vpn['Ping'][workbook.vpn['IP']==ip]='UP'
    else:
        print(f"DOWN {ip} Ping Unsuccessful")
        workbook.vpn['Ping'][workbook.vpn['IP']==ip]='DOWN'
def vpn_lookup(ip):
    response = os.popen("nslookup "+ ip+" | awk 'NR==5{print $4}'").read()
    if response == '\n':
        print(f"{ip} No DNS")
        workbook.vpn['CNAME'][workbook.vpn['IP']==ip]='None'
    else:
        print(f"{ip} {response}")
        response = response.rstrip()
        workbook.vpn['CNAME'][workbook.vpn['IP']==ip]=response
        
def blank_ping(ip):
    response = os.popen(f"ping -c 1 {ip}").read()
    if "Received = 4" in response:
        print(f"UP {ip} Ping Successful")
        workbook.blank['Ping'][workbook.blank['IP']==ip]='UP'
    else:
        print(f"DOWN {ip} Ping Unsuccessful")
        workbook.blank['Ping'][workbook.blank['IP']==ip]='DOWN'
def blank_lookup(ip):
    response = os.popen("nslookup "+ ip+" | awk 'NR==5{print $4}'").read()
    if response == '\n':
        print(f"{ip} No DNS")
        workbook.blank['CNAME'][workbook.blank['IP']==ip]='None'
    else:
        print(f"{ip} {response}")
        response = response.rstrip()
        workbook.blank['CNAME'][workbook.blank['IP']==ip]=response
        
def console_ping(ip):
    response = os.popen(f"ping -c 1 {ip}").read()
    if "Received = 4" in response:
        print(f"UP {ip} Ping Successful")
        workbook.console['Ping'][workbook.console['IP']==ip]='UP'
    else:
        print(f"DOWN {ip} Ping Unsuccessful")
        workbook.console['Ping'][workbook.console['IP']==ip]='DOWN'
def console_lookup(ip):
    response = os.popen("nslookup "+ ip+" | awk 'NR==5{print $4}'").read()
    if response == '\n':
        print(f"{ip} No DNS")
        workbook.console['CNAME'][workbook.console['IP']==ip]='None'
    else:
        print(f"{ip} {response}")
        response = response.rstrip()
        workbook.console['CNAME'][workbook.console['IP']==ip]=response
        
def interface_ping(ip):
    response = os.popen(f"ping -c 1 {ip}").read()
    if "Received = 4" in response:
        print(f"UP {ip} Ping Successful")
        workbook.interface['Ping'][workbook.interface['IP']==ip]='UP'
    else:
        print(f"DOWN {ip} Ping Unsuccessful")
        workbook.interface['Ping'][workbook.interface['IP']==ip]='DOWN'
def interface_lookup(ip):
    response = os.popen("nslookup "+ ip+" | awk 'NR==5{print $4}'").read()
    if response == '\n':
        print(f"{ip} No DNS")
        workbook.interface['CNAME'][workbook.interface['IP']==ip]='None'
    else:
        print(f"{ip} {response}")
        response = response.rstrip()
        workbook.interface['CNAME'][workbook.interface['IP']==ip]=response

def nouse_ping(ip):
    response = os.popen(f"ping -c 1 {ip}").read()
    if "Received = 4" in response:
        print(f"UP {ip} Ping Successful")
        workbook.nouse['Ping'][workbook.nouse['IP']==ip]='UP'
    else:
        print(f"DOWN {ip} Ping Unsuccessful")
        workbook.nouse['Ping'][workbook.nouse['IP']==ip]='DOWN'
def nouse_lookup(ip):
    response = os.popen("nslookup "+ ip+" | awk 'NR==5{print $4}'").read()
    if response == '\n':
        print(f"{ip} No DNS")
        workbook.nouse['CNAME'][workbook.nouse['IP']==ip]='None'
    else:
        print(f"{ip} {response}")
        response = response.rstrip()
        workbook.nouse['CNAME'][workbook.nouse['IP']==ip]=response
                
def vm_ping(ip):
    response = os.popen(f"ping -c 1 {ip}").read()
    if "Received = 4" in response:
        print(f"UP {ip} Ping Successful")
        workbook.vm['Ping'][workbook.vm['IP']==ip]='UP'
    else:
        print(f"DOWN {ip} Ping Unsuccessful")
        workbook.vm['Ping'][workbook.vm['IP']==ip]='DOWN'
def vm_lookup(ip):
    response = os.popen("nslookup "+ ip+" | awk 'NR==5{print $4}'").read()
    if response == '\n':
        print(f"{ip} No DNS")
        workbook.vm['CNAME'][workbook.vm['IP']==ip]='None'
    else:
        print(f"{ip} {response}")
        response = response.rstrip()
        workbook.vm['CNAME'][workbook.vm['IP']==ip]=response
        
def available_ping(ip):
    response = os.popen(f"ping -c 1 {ip}").read()
    if "Received = 4" in response:
        print(f"UP {ip} Ping Successful")
        workbook.available['Ping'][workbook.available['IP']==ip]='UP'
    else:
        print(f"DOWN {ip} Ping Unsuccessful")
        workbook.available['Ping'][workbook.available['IP']==ip]='DOWN'
def available_lookup(ip):
    response = os.popen("nslookup "+ ip+" | awk 'NR==5{print $4}'").read()
    if response == '\n':
        print(f"{ip} No DNS")
        workbook.available['CNAME'][workbook.available['IP']==ip]='None'
    else:
        print(f"{ip} {response}")
        response = response.rstrip()
        workbook.available['CNAME'][workbook.available['IP']==ip]=response
                
start = time.perf_counter()
workbook = pd.ExcelFile('not_in_cmdb_Data.xlsx')


        

workbook.main = pd.read_excel(workbook,sheet_name='Main')
workbook.main['Ping']=''
workbook.main['CNAME']=''

workbook.duplicate_DNS = pd.read_excel(workbook,sheet_name='Duplicate DNS')
workbook.duplicate_DNS['Ping']=''
workbook.duplicate_DNS['CNAME']=''

workbook.duplicate_IP = pd.read_excel(workbook,sheet_name='Duplicate IP')
workbook.duplicate_IP['Ping']=''
workbook.duplicate_IP['CNAME']=''


workbook.ls = pd.read_excel(workbook,sheet_name='ls')
workbook.ls['Ping']=''
workbook.ls['CNAME']=''

workbook.cy = pd.read_excel(workbook,sheet_name='cy')
workbook.cy['Ping']=''
workbook.cy['CNAME']=''

workbook.pdu = pd.read_excel(workbook,sheet_name='PDU')
workbook.pdu['Ping']=''
workbook.pdu['CNAME']=''

workbook.pub = pd.read_excel(workbook,sheet_name='Public Facing')
workbook.pub['Ping']=''
workbook.pub['CNAME']=''

workbook.bo1 = pd.read_excel(workbook,sheet_name='BO1')
workbook.bo1['Ping']=''
workbook.bo1['CNAME']=''

workbook.agent = pd.read_excel(workbook,sheet_name='AGENT')
workbook.agent['Ping']=''
workbook.agent['CNAME']=''

workbook.bronto = pd.read_excel(workbook,sheet_name='Bronto')
workbook.bronto['Ping']=''
workbook.bronto['CNAME']=''


workbook.openair = pd.read_excel(workbook,sheet_name='OpenAir')
workbook.openair['Ping']=''
workbook.openair['CNAME']=''

workbook.oci = pd.read_excel(workbook,sheet_name='OCI')
workbook.oci['Ping']=''
workbook.oci['CNAME']=''


workbook.payroll = pd.read_excel(workbook,sheet_name='Payroll')
workbook.payroll['Ping']=''
workbook.payroll['CNAME']=''

workbook.vpn = pd.read_excel(workbook,sheet_name='VPN')
workbook.vpn['Ping']=''
workbook.vpn['CNAME']=''


workbook.blank = pd.read_excel(workbook,sheet_name='None')
workbook.blank['Ping']=''
workbook.blank['CNAME']=''

workbook.console = pd.read_excel(workbook,sheet_name='Console')
workbook.console['Ping']=''
workbook.console['CNAME']=''

workbook.interface = pd.read_excel(workbook,sheet_name='Interface')
workbook.interface ['Ping']=''
workbook.interface['CNAME']=''

workbook.nouse = pd.read_excel(workbook,sheet_name='NoUse')
workbook.nouse['Ping']=''
workbook.nouse['CNAME']=''

workbook.vm = pd.read_excel(workbook,sheet_name='VM')
workbook.vm['Ping']=''
workbook.vm['CNAME']=''

workbook.available = pd.read_excel(workbook,sheet_name='Available')
workbook.available['Ping']=''
workbook.available['CNAME']=''

with concurrent.futures.ThreadPoolExecutor() as executor:
    executor.map(main_ping,workbook.main['IP'])
    executor.map(main_lookup,workbook.main['IP'])
    executor.map(ls_ping,workbook.ls['IP'])
    executor.map(ls_lookup,workbook.ls['IP'])
    executor.map(cy_ping,workbook.cy['IP'])
    executor.map(cy_lookup,workbook.cy['IP'])
    executor.map(pdu_ping,workbook.pdu['IP'])
    executor.map(pdu_lookup,workbook.pdu['IP'])
    executor.map(pub_ping,workbook.pub['IP'])
    executor.map(pub_lookup,workbook.pub['IP'])
    executor.map(bo1_ping,workbook.bo1['IP'])
    executor.map(bo1_lookup,workbook.bo1['IP'])
    executor.map(agent_ping,workbook.agent['IP'])
    executor.map(agent_lookup,workbook.agent['IP'])
    executor.map(bronto_ping,workbook.bronto['IP'])
    executor.map(bronto_lookup,workbook.bronto['IP'])
    executor.map(openair_ping,workbook.openair['IP'])
    executor.map(openair_lookup,workbook.openair['IP'])
    executor.map(oci_ping,workbook.oci['IP'])
    executor.map(oci_lookup,workbook.oci['IP'])
    executor.map(payroll_ping,workbook.payroll['IP'])
    executor.map(payroll_lookup,workbook.payroll['IP'])
    executor.map(vpn_ping,workbook.vpn['IP'])
    executor.map(vpn_lookup,workbook.vpn['IP'])
    executor.map(blank_ping,workbook.blank['IP'])
    executor.map(blank_lookup,workbook.blank['IP'])
    executor.map(console_ping,workbook.console['IP'])
    executor.map(console_lookup,workbook.console['IP'])
    executor.map(interface_ping,workbook.interface['IP'])
    executor.map(interface_lookup,workbook.interface['IP'])
    executor.map(nouse_ping,workbook.nouse['IP'])
    executor.map(nouse_lookup,workbook.nouse['IP'])
    executor.map(vm_ping,workbook.vm['IP'])
    executor.map(vm_lookup,workbook.vm['IP'])
    executor.map(available_ping,workbook.available['IP'])
    executor.map(available_lookup,workbook.available['IP'])
    
    
with pd.ExcelWriter('not_in_CMDB_ping_data.xlsx') as writer:
    workbook.main.to_excel(writer,sheet_name='Main', index=False)
    workbook.duplicate_DNS.to_excel(writer,sheet_name='Duplicate DNS', index=False)
    workbook.duplicate_IP.to_excel(writer,sheet_name='Duplicate IP', index=False)
    workbook.ls.to_excel(writer,sheet_name='ls', index=False)
    workbook.cy.to_excel(writer,sheet_name='cy', index=False)
    workbook.pdu.to_excel(writer,sheet_name='PDU', index=False)
    workbook.pub.to_excel(writer,sheet_name='Public Facing', index=False)
    workbook.bo1.to_excel(writer,sheet_name='BO1', index=False)
    workbook.agent.to_excel(writer,sheet_name='AGENT', index=False)
    workbook.bronto.to_excel(writer,sheet_name='Bronto', index=False)
    workbook.payroll.to_excel(writer,sheet_name='Payroll', index=False)
    workbook.oci.to_excel(writer,sheet_name='OCI', index=False)
    workbook.openair.to_excel(writer,sheet_name='OpenAir', index=False)
    workbook.vpn.to_excel(writer,sheet_name='VPN', index=False)
    workbook.blank.to_excel(writer,sheet_name='None', index=False)
    workbook.console.to_excel(writer,sheet_name='Console', index=False)
    workbook.interface.to_excel(writer,sheet_name='Interface', index=False)
    workbook.nouse.to_excel(writer,sheet_name='NoUse', index=False)
    workbook.vm.to_excel(writer,sheet_name='VM', index=False)
    workbook.available.to_excel(writer,sheet_name='Available', index=False)
    
stop = time.perf_counter()

print("It took "+str(round((stop-start)/60,2))+" minutes for the script to complete")