#import modules
import pandas as pd
import os.path
import shutil, time
import ipaddress

#ask for cloud file name
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


print("Merging data")
merg = cloud.append(pcp)
print("Data count after merge: " + str(len(merg)))
init_count = len(merg)
time.sleep(2)
#separating duplicates duplicates

#separating none DNS
print("Removing data containing none in DNS column")
blank = merg[merg.DNS.str.contains("None|none",na=False)]

print("Removing duplicated IP")
merg_duplicate_IP = merg[merg.duplicated('IP',keep=False)]
print("Removing duplicated DNS")
merg_duplicate_DNS = merg[merg.duplicated('DNS',keep=False)]

#clean data from duplicates
print("Pre cleanup count: " +str(len(merg)))
merg = merg[~merg.duplicated('IP',keep=False)]
print("removed duplicated IP count: " + str(len(merg)))
merg = merg[~merg.duplicated('DNS',keep=False)]
print("removed duplicated DNS count: " + str(len(merg)))
time.sleep(2)


#separating Bronto tags and DNS
print("Removing Bronto")
#separating data that has bronto in DNS
bronto = merg[merg.DNS.str.contains("Bronto|bronto",na=False)]
#separating data that has bronto in tags
bronto = bronto[bronto.Tags.str.contains("Bronto|bronto",na=False)]

print("Pre cleanup count: " +str(len(merg)))
merg = merg[~merg.DNS.str.contains("Bronto|bronto",na=False)]
print("removed data that contains bronto in DNS count: " +str(len(merg)))
merg = merg[~merg.Tags.str.contains("Bronto|bronto",na=False)]
print("removed data that contains bronto in TAGS count: " +str(len(merg)))
time.sleep(2)

#separating OpenAir Tags
print("Removing data containing OpenAir in Tags column")
#separating data that has bronto in DNS
openair = merg[merg.Tags.str.contains("openair|OpenAir",na=False)]

print("Pre cleanup count: " +str(len(merg)))
merg = merg[~merg.Tags.str.contains("openair|OpenAir",na=False)]
print("removed data that contains OpenAir in TAGS count: " +str(len(merg)))
time.sleep(2)

#separating OCI tags and DNS
print("Removing OCI")
#separating data that has oci in DNS
oci = merg[merg.DNS.str.contains("OCI|oci",na=False)]
#separating data that has bronto in tags
oci = oci[oci.Tags.str.contains("OCI|oci",na=False)]

print("Pre cleanup count: " +str(len(merg)))
merg = merg[~merg.DNS.str.contains("OCI|oci",na=False)]
print("removed data that contains OCI in DNS count: " +str(len(merg)))
merg = merg[~merg.Tags.str.contains("OCI|oci",na=False)]
print("removed data that contains OCI in TAGS count: " +str(len(merg)))
time.sleep(2)

#separating Payroll Tags
print("Removing data containing Payroll in Tags column")
#separating data that has bronto in DNS
payroll = merg[merg.Tags.str.contains("payroll|Payroll",na=False)]

print("Pre cleanup count: " +str(len(merg)))
merg = merg[~merg.Tags.str.contains("payroll|Payroll",na=False)]
print("removed data that contains Payroll in TAGS count: " +str(len(merg)))
time.sleep(2)

#separating lcms tags and DNS
print("Removing LCMS")
#separating data that has oci in DNS
lcms = merg[merg.DNS.str.contains("LCMS|lcms",na=False)]
#separating data that has bronto in tags
lcms = lcms[lcms.Tags.str.contains("LCMS|lcms",na=False)]

print("Pre cleanup count: " +str(len(merg)))
merg = merg[~merg.DNS.str.contains("LCMS|lcms",na=False)]
print("removed data that contains LCMS in DNS count: " +str(len(merg)))
merg = merg[~merg.Tags.str.contains("LCMS|lcms",na=False)]
print("removed data that contains LCMS in TAGS count: " +str(len(merg)))
time.sleep(2)

#separating vpns DNS
print("Removing data containing VPN in DNS column")
vpn = merg[merg.DNS.str.contains("vpn|ovpn",na=False)]
vpn = merg[merg.Tags.str.contains("vpn|ovpn",na=False)]

print("Pre cleanup count: " +str(len(merg)))
merg = merg[~merg.DNS.str.contains("vpn|ovpn",na=False)]
print("removed data that contains VPN in DNS count: " +str(len(merg)))
merg = merg[~merg.Tags.str.contains("vpn|ovpn",na=False)]
print("removed data that contains VPN in TAGS count: " +str(len(merg)))
time.sleep(2)


print("Pre cleanup count: " +str(len(merg)))
merg = merg[~merg.DNS.str.contains("None|none",na=False)]
print("removed data that contains None in DNS count: " +str(len(merg)))
time.sleep(2)

#separating console DNS
print("Removing data containing console in DNS column")
console = merg[merg.DNS.str.contains("console",na=False)]
console = console[console.Tags.str.contains("console",na=False)]


print("Pre cleanup count: " +str(len(merg)))
merg = merg[~merg.DNS.str.contains("console",na=False)]
print("removed data that contains console in DNS count: " +str(len(merg)))
merg = merg[~merg.Tags.str.contains("console",na=False)]
print("removed data that contains console in TAGS count: " +str(len(merg)))
time.sleep(2)


#separating Interfaces DNS
print("Removing data containing interface in DNS column")
interface = merg[merg.DNS.str.contains("interface",na=False)]


print("Pre cleanup count: " +str(len(merg)))
merg = merg[~merg.DNS.str.contains("interface",na=False)]
print("removed data that contains interface in DNS count: " +str(len(merg)))
time.sleep(2)

#separating nouse DNS
print("Removing data containing interface in DNS column")
nouse = merg[merg.DNS.str.contains("nouse",na=False)]


print("Pre cleanup count: " +str(len(merg)))
merg = merg[~merg.DNS.str.contains("nouse",na=False)]
print("removed data that contains nouse in DNS count: " +str(len(merg)))
time.sleep(2)

#separating vm DNS
print("Removing data containing snap|\.f\.|\.m\. in DNS column")
vm = merg[merg.DNS.str.contains("snap|\.f\.|\.m\.",na=False)]

print("Pre cleanup count: " +str(len(merg)))
merg = merg[~merg.DNS.str.contains("snap|\.f\.|\.m\.",na=False)]
print("removed data that contains snap|\.f\.|\.m\. in DNS count: " +str(len(merg)))
time.sleep(2)


#separating AVAILABLE DNS
print("Removing data containing available in DNS column")
available = merg[merg.DNS.str.contains("available",na=False)]


print("Pre cleanup count: " +str(len(merg)))
merg = merg[~merg.DNS.str.contains("available",na=False)]
print("removed data that contains available in DNS count: " +str(len(merg)))
time.sleep(2)


ip_list = ['10.16.212.0/22','10.18.204.0/22','10.18.212.0/22','10.9.28.0/22','10.9.32.0/22','10.3.204.0/22','10.3.216.0/22','10.2.212.0/22','10.2.216.0/22','10.1.240.0/22','10.1.168.0/21']
print("Pre cleanup count: " +str(len(merg)))
for ip in ipaddress.IPv4Network('10.16.204.0/22'):
        print(ip)
        ip_new = merg[merg.IP.str.contains(str(ip),na=False)]
        merg = merg[~merg.IP.str.contains(str(ip),na=False)]
for ip_range in ip_list:
    for ip in ipaddress.IPv4Network(ip_range):
        print(ip)
        ip_new = ip_new[ip_new.IP.str.contains(str(ip),na=False)]
        merg = merg[~merg.IP.str.contains(str(ip),na=False)]
print("Data that contains VPN IP in IP Column count: " +str(len(ip_new)))
print("removed data that contains VPN IP in IP Column count: " +str(len(merg)))
time.sleep(2)

print("exporting File")
file_name = input("Enter filename ex: New_Data (don't include file extensions)")
with pd.ExcelWriter(file_name+'.xlsx') as writer:
    merg.to_excel(writer,sheet_name ='Main',index=False)
    merg_duplicate_IP.to_excel(writer,sheet_name='Duplicate IP',index=False)
    merg_duplicate_DNS.to_excel(writer,sheet_name='Duplicate DNS',index=False)
    bronto.to_excel(writer,sheet_name='Bronto',index=False)
    openair.to_excel(writer,sheet_name='OpenAir',index=False)
    oci.to_excel(writer,sheet_name='OCI',index=False)
    payroll.to_excel(writer,sheet_name='Payroll',index=False)
    lcms.to_excel(writer,sheet_name='LCMS',index=False)
    vpn.to_excel(writer,sheet_name='VPN',index=False)
    blank.to_excel(writer,sheet_name='None',index=False)
    console.to_excel(writer,sheet_name='Console',index=False)
    interface.to_excel(writer,sheet_name='Interface',index=False)
    nouse.to_excel(writer,sheet_name='NoUse',index=False)
    vm.to_excel(writer,sheet_name='VM',index=False)
    available.to_excel(writer,sheet_name='Available',index=False)
    #VPN IP ranges are not in the list
    # ip_new.to_excel(writer,sheet_name='VPN_IP_Range',index=False)
stop = time.perf_counter()
print(file_name+" has been created")


print("It took "+str(round((stop-start)/60,2))+" minutes for the script to complete")




