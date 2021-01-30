import pandas as pd
import time
start = time.perf_counter()

workbook = pd.ExcelFile('CMDB_queried_data_latest_2_0.xlsx')
data = pd.read_excel(workbook)

ip_in_cmdb = data[data['IP in CMDB'].str.contains('IP in CMDB',na=False)]
ip_not_in_cmdb = data[~data['IP in CMDB'].str.contains('IP in CMDB',na=False)]

dns_in_cmdb = ip_not_in_cmdb[ip_not_in_cmdb['DNS in CMDB'].str.contains('DNS in CMDB', na=False)]
dns_not_in_cmdb = ip_not_in_cmdb[~ip_not_in_cmdb['DNS in CMDB'].str.contains('DNS in CMDB', na=False)]

in_cmdb = ip_in_cmdb.append(dns_in_cmdb)
not_in_cmdb = dns_not_in_cmdb


#None
blank_in = in_cmdb[in_cmdb.DNS.str.contains("None|none",na=False)]
blank_not = not_in_cmdb[not_in_cmdb.DNS.str.contains("None|none",na=False)]


in_cmdb = in_cmdb[~in_cmdb.DNS.str.contains("None|none",na=False)]
not_in_cmdb = not_in_cmdb[~not_in_cmdb.DNS.str.contains("None|none",na=False)]

#Duplicate
in_cmdb_duplicate_ip = in_cmdb[in_cmdb.duplicated('IP',keep=False)]
in_cmdb_duplicate_dns = in_cmdb[in_cmdb.duplicated('DNS',keep=False)]

not_in_cmdb_duplicate_ip = not_in_cmdb[not_in_cmdb.duplicated('IP',keep=False)]
not_in_cmdb_duplicate_dns = not_in_cmdb[not_in_cmdb.duplicated('DNS',keep=False)]

in_cmdb = in_cmdb[~in_cmdb.duplicated('IP',keep=False)]
in_cmdb = in_cmdb[~in_cmdb.duplicated('DNS',keep=False)]

not_in_cmdb = not_in_cmdb[~not_in_cmdb.duplicated('IP',keep=False)]
not_in_cmdb = not_in_cmdb[~not_in_cmdb.duplicated('DNS',keep=False)]


#Tracking
agent_in = in_cmdb[in_cmdb['Tracking Method'].str.contains("AGENT", na=False)]
agent_out = not_in_cmdb[not_in_cmdb['Tracking Method'].str.contains("AGENT", na=False)]

in_cmdb= in_cmdb[~in_cmdb['Tracking Method'].str.contains("AGENT", na=False)]
not_in_cmdb = not_in_cmdb[~not_in_cmdb['Tracking Method'].str.contains("AGENT", na=False)]

#LS/Network Switch
ls_in = in_cmdb[in_cmdb.DNS.str.contains("-ls",na=False)]
cy_in = in_cmdb[in_cmdb.DNS.str.contains("cy",na=False)]

ls_out = not_in_cmdb[not_in_cmdb.DNS.str.contains("-ls",na=False)]
cy_out = not_in_cmdb[not_in_cmdb.DNS.str.contains("-cy",na=False)]

in_cmdb = in_cmdb[~in_cmdb.DNS.str.contains("-ls",na=False)]
in_cmdb = in_cmdb[~in_cmdb.DNS.str.contains("cy",na=False)]

not_in_cmdb = not_in_cmdb[~not_in_cmdb.DNS.str.contains("-ls",na=False)]
not_in_cmdb = not_in_cmdb[~not_in_cmdb.DNS.str.contains("-cy",na=False)]

#pdu
pdu_in = in_cmdb[in_cmdb['DNS'].str.contains("pdu", na=False)]
pdu_out = not_in_cmdb[not_in_cmdb['DNS'].str.contains("pdu", na=False)]

in_cmdb= in_cmdb[~in_cmdb['DNS'].str.contains("pdu", na=False)]
not_in_cmdb = not_in_cmdb[~not_in_cmdb['DNS'].str.contains("pdu", na=False)]

#Internet facing

public_in = in_cmdb[in_cmdb.Tags.str.contains("Internet Facing Assets",na=False)]

public_out = not_in_cmdb[not_in_cmdb.Tags.str.contains("Internet Facing Assets",na=False)]

in_cmdb = in_cmdb[~in_cmdb.Tags.str.contains("Internet Facing Assets",na=False)]

not_in_cmdb = not_in_cmdb[~not_in_cmdb.Tags.str.contains("Internet Facing Assets",na=False)]

#Internet facing

bo1_in = in_cmdb[in_cmdb.DNS.str.contains("BO1|bo1|b01",na=False)]
bo1_in = bo1_in[bo1_in.Tags.str.contains("BO1|bo1|b01",na=False)]

bo1_out = not_in_cmdb[not_in_cmdb.DNS.str.contains("BO1|bo1|b01",na=False)]
bo1_out = bo1_out[bo1_out.Tags.str.contains("BO1|bo1|b01",na=False)]

in_cmdb = in_cmdb[~in_cmdb.DNS.str.contains("BO1|bo1|b01",na=False)]
in_cmdb = in_cmdb[~in_cmdb.Tags.str.contains("BO1|bo1|b01",na=False)]

not_in_cmdb = not_in_cmdb[~not_in_cmdb.DNS.str.contains("BO1|bo1|b01",na=False)]
not_in_cmdb = not_in_cmdb[~not_in_cmdb.Tags.str.contains("BO1|bo1|b01",na=False)]


#Bronto
bronto_in = in_cmdb[in_cmdb.DNS.str.contains("Bronto|bronto",na=False)]
bronto_in = bronto_in[bronto_in.Tags.str.contains("Bronto|bronto",na=False)]

bronto_not = not_in_cmdb[not_in_cmdb.DNS.str.contains("Bronto|bronto",na=False)]
bronto_not = bronto_not[bronto_not.Tags.str.contains("Bronto|bronto",na=False)]

in_cmdb = in_cmdb[~in_cmdb.DNS.str.contains("Bronto|bronto",na=False)]
in_cmdb = in_cmdb[~in_cmdb.Tags.str.contains("Bronto|bronto",na=False)]

not_in_cmdb = not_in_cmdb[~not_in_cmdb.DNS.str.contains("Bronto|bronto",na=False)]
not_in_cmdb = not_in_cmdb[~not_in_cmdb.Tags.str.contains("Bronto|bronto",na=False)]

#OpenAir
openair_in = in_cmdb[in_cmdb.Tags.str.contains("openair|OpenAir",na=False)]
openair_not = not_in_cmdb[not_in_cmdb.Tags.str.contains("openair|OpenAir",na=False)]

in_cmdb = in_cmdb[~in_cmdb.Tags.str.contains("openair|OpenAir",na=False)]
not_in_cmdb = not_in_cmdb[~not_in_cmdb.Tags.str.contains("openair|OpenAir",na=False)]

#OCI

oci_in = in_cmdb[in_cmdb.DNS.str.contains("OCI|oci",na=False)]
oci_in = oci_in[oci_in.Tags.str.contains("OCI|oci",na=False)]

oci_not = not_in_cmdb[not_in_cmdb.DNS.str.contains("OCI|oci",na=False)]
oci_not = oci_not[oci_not.Tags.str.contains("OCI|oci",na=False)]


in_cmdb = in_cmdb[~in_cmdb.DNS.str.contains("OCI|oci",na=False)]
in_cmdb = in_cmdb[~in_cmdb.Tags.str.contains("OCI|oci",na=False)]

not_in_cmdb = not_in_cmdb[~not_in_cmdb.DNS.str.contains("OCI|oci",na=False)]
not_in_cmdb = not_in_cmdb[~not_in_cmdb.Tags.str.contains("OCI|oci",na=False)]


#Payroll

payroll_in = in_cmdb[in_cmdb.Tags.str.contains("payroll|Payroll",na=False)]

payroll_not = not_in_cmdb[not_in_cmdb.Tags.str.contains("payroll|Payroll",na=False)]


in_cmdb = in_cmdb[~in_cmdb.Tags.str.contains("payroll|Payroll",na=False)]

not_in_cmdb = not_in_cmdb[~not_in_cmdb.Tags.str.contains("payroll|Payroll",na=False)]


#VPN

vpn_in = in_cmdb[in_cmdb.DNS.str.contains("vpn|ovpn",na=False)]
vpn_in = in_cmdb[in_cmdb.Tags.str.contains("vpn|ovpn",na=False)]

vpn_not = not_in_cmdb[not_in_cmdb.DNS.str.contains("vpn|ovpn",na=False)]
vpn_not = not_in_cmdb[not_in_cmdb.Tags.str.contains("vpn|ovpn",na=False)]

in_cmdb = in_cmdb[~in_cmdb.DNS.str.contains("vpn|ovpn",na=False)]
in_cmdb = in_cmdb[~in_cmdb.Tags.str.contains("vpn|ovpn",na=False)]

not_in_cmdb = not_in_cmdb[~not_in_cmdb.DNS.str.contains("vpn|ovpn",na=False)]
not_in_cmdb = not_in_cmdb[~not_in_cmdb.Tags.str.contains("vpn|ovpn",na=False)]

#console
console_in = in_cmdb[in_cmdb.DNS.str.contains("console",na=False)]
console_in = console_in[console_in.Tags.str.contains("console",na=False)]

console_not = not_in_cmdb[not_in_cmdb.DNS.str.contains("console",na=False)]
console_not = console_not[console_not.Tags.str.contains("console",na=False)]

in_cmdb = in_cmdb[~in_cmdb.DNS.str.contains("console",na=False)]
in_cmdb = in_cmdb[~in_cmdb.Tags.str.contains("console",na=False)]

not_in_cmdb = not_in_cmdb[~not_in_cmdb.DNS.str.contains("console",na=False)]
not_in_cmdb = not_in_cmdb[~not_in_cmdb.Tags.str.contains("console",na=False)]

#interface
interface_in = in_cmdb[in_cmdb.DNS.str.contains("interface",na=False)]

interface_not = not_in_cmdb[not_in_cmdb.DNS.str.contains("interface",na=False)]

in_cmdb = in_cmdb[~in_cmdb.DNS.str.contains("interface",na=False)]

not_in_cmdb = not_in_cmdb[~not_in_cmdb.DNS.str.contains("interface",na=False)]

#NoUse

nouse_in = in_cmdb[in_cmdb.DNS.str.contains("nouse",na=False)]

nouse_not = not_in_cmdb[not_in_cmdb.DNS.str.contains("nouse",na=False)]

in_cmdb = in_cmdb[~in_cmdb.DNS.str.contains("nouse",na=False)]

not_in_cmdb = not_in_cmdb[~not_in_cmdb.DNS.str.contains("nouse",na=False)]

#VM
vm_in = in_cmdb[in_cmdb.DNS.str.contains("snap|\.f\.|\.m\.",na=False)]

vm_not = not_in_cmdb[not_in_cmdb.DNS.str.contains("snap|\.f\.|\.m\.",na=False)]


in_cmdb = in_cmdb[~in_cmdb.DNS.str.contains("snap|\.f\.|\.m\.",na=False)]

not_in_cmdb = not_in_cmdb[~not_in_cmdb.DNS.str.contains("snap|\.f\.|\.m\.",na=False)]


#Available
available_in = in_cmdb[in_cmdb.DNS.str.contains("available",na=False)]

available_not = not_in_cmdb[not_in_cmdb.DNS.str.contains("available",na=False)]

in_cmdb = in_cmdb[~in_cmdb.DNS.str.contains("available",na=False)]

not_in_cmdb = not_in_cmdb[~not_in_cmdb.DNS.str.contains("available",na=False)]

print("exporting Files")
in_cmdb_file_name = 'in_cmdb_Data(1).xlsx'
writer = pd.ExcelWriter(in_cmdb_file_name)
in_cmdb.to_excel(writer,sheet_name ='Main',index=False)
in_cmdb_duplicate_ip.to_excel(writer,sheet_name='Duplicate IP',index=False)
in_cmdb_duplicate_dns.to_excel(writer,sheet_name='Duplicate DNS',index=False)
ls_in.to_excel(writer,sheet_name='ls', index=False)
cy_in.to_excel(writer,sheet_name='cy', index=False)
pdu_in.to_excel(writer,sheet_name='PDU', index=False)
public_in.to_excel(writer,sheet_name='Public Facing', index=False)
bo1_in.to_excel(writer,sheet_name='BO1', index=False)
agent_in.to_excel(writer,sheet_name='AGENT',index=False)
bronto_in.to_excel(writer,sheet_name='Bronto',index=False)
openair_in.to_excel(writer,sheet_name='OpenAir',index=False)
oci_in.to_excel(writer,sheet_name='OCI',index=False)
payroll_in.to_excel(writer,sheet_name='Payroll',index=False)
vpn_in.to_excel(writer,sheet_name='VPN',index=False)
blank_in.to_excel(writer,sheet_name='None',index=False)
console_in.to_excel(writer,sheet_name='Console',index=False)
interface_in.to_excel(writer,sheet_name='Interface',index=False)
nouse_in.to_excel(writer,sheet_name='NoUse',index=False)
vm_in.to_excel(writer,sheet_name='VM',index=False)
available_in.to_excel(writer,sheet_name='Available',index=False)
writer.save()

not_in_cmdb_file_name = 'not_in_cmdb_Data(1).xlsx'
writer1 =  pd.ExcelWriter(not_in_cmdb_file_name)
not_in_cmdb.to_excel(writer1,sheet_name ='Main',index=False)
not_in_cmdb_duplicate_ip.to_excel(writer1,sheet_name='Duplicate IP',index=False)
not_in_cmdb_duplicate_dns.to_excel(writer1,sheet_name='Duplicate DNS',index=False)
ls_out.to_excel(writer1,sheet_name='ls', index=False)
cy_out.to_excel(writer1,sheet_name='cy', index=False)
pdu_out.to_excel(writer1,sheet_name='PDU', index=False)
public_out.to_excel(writer1,sheet_name='Public Facing', index=False)
bo1_out.to_excel(writer1,sheet_name='BO1', index=False)
agent_out.to_excel(writer1,sheet_name='AGENT',index=False)
bronto_not.to_excel(writer1,sheet_name='Bronto',index=False)
openair_not.to_excel(writer1,sheet_name='OpenAir',index=False)
oci_not.to_excel(writer1,sheet_name='OCI',index=False)
payroll_not.to_excel(writer1,sheet_name='Payroll',index=False)
vpn_not.to_excel(writer1,sheet_name='VPN',index=False)
blank_not.to_excel(writer1,sheet_name='None',index=False)
console_not.to_excel(writer1,sheet_name='Console',index=False)
interface_not.to_excel(writer1,sheet_name='Interface',index=False)
nouse_not.to_excel(writer1,sheet_name='NoUse',index=False)
vm_not.to_excel(writer1,sheet_name='VM',index=False)
available_not.to_excel(writer1,sheet_name='Available',index=False)
writer1.save()
    
stop = time.perf_counter()

print("It took "+str(round((stop-start)/60,2))+" minutes for the script to complete")