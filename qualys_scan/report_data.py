    
import pandas as pd
import time

start = time.perf_counter()
print("Loading Workbook")

workbook = pd.ExcelFile('CMDB_queried_data_latest.xlsx')
workbook.main = pd.read_excel(workbook,sheet_name='Main')
workbook.duplicate_IP = pd.read_excel(workbook,sheet_name='Duplicate IP')
workbook.duplicate_DNS = pd.read_excel(workbook,sheet_name='Duplicate DNS')
workbook.bronto = pd.read_excel(workbook,sheet_name='Bronto')
workbook.openair = pd.read_excel(workbook,sheet_name='OpenAir')
workbook.oci = pd.read_excel(workbook,sheet_name='OCI')
workbook.payroll = pd.read_excel(workbook,sheet_name='OCI')
workbook.vpn = pd.read_excel(workbook,sheet_name='VPN')
workbook.blank = pd.read_excel(workbook,sheet_name='None')
workbook.console = pd.read_excel(workbook,sheet_name='Console')
workbook.interface = pd.read_excel(workbook,sheet_name='Interface')
workbook.nouse = pd.read_excel(workbook,sheet_name='NoUse')
workbook.vm = pd.read_excel(workbook,sheet_name='VM')
workbook.available = pd.read_excel(workbook,sheet_name='Available')


workbook.main['IP in CMDB Count']=''
workbook.main['IP not in CMDB Count']=''
workbook.main['DNS not in CMDB Count']= ''
workbook.main['DNS in CMDB Count']=''
workbook.main['DNS in CMDB + Count']=''

workbook.main['IP in CMDB Count'][0]= len(workbook.main[workbook.main['IP in CMDB']== 'IP in CMDB'])
workbook.main['IP not in CMDB Count'][0]= len(workbook.main[workbook.main['IP in CMDB']== 'IP NOT in CMDB'])
workbook.main['DNS in CMDB Count'][0]= len(workbook.main[workbook.main['DNS in CMDB']== 'DNS in CMDB'])
workbook.main['DNS not in CMDB Count'][0]= len(workbook.main[workbook.main['DNS in CMDB']== 'DNS NOT in CMDB'])
workbook.main['DNS in CMDB + Count'][0]= len(workbook.main[workbook.main['DNS in CMDB']== 'DNS in CMDB + '])

workbook.bronto['IP in CMDB Count']=''
workbook.bronto['IP not in CMDB Count']=''
workbook.bronto['DNS not in CMDB Count']= ''
workbook.bronto['DNS in CMDB Count']=''
workbook.bronto['DNS in CMDB + Count']=''

workbook.bronto['IP in CMDB Count'][0]= len(workbook.bronto[workbook.bronto['IP in CMDB']== 'IP in CMDB'])
workbook.bronto['IP not in CMDB Count'][0]= len(workbook.bronto[workbook.bronto['IP in CMDB']== 'IP NOT in CMDB'])
workbook.bronto['DNS in CMDB Count'][0]= len(workbook.bronto[workbook.bronto['DNS in CMDB']== 'DNS in CMDB'])
workbook.bronto['DNS not in CMDB Count'][0]= len(workbook.bronto[workbook.bronto['DNS in CMDB']== 'DNS NOT in CMDB'])
workbook.bronto['DNS in CMDB + Count'][0]= len(workbook.bronto[workbook.bronto['DNS in CMDB']== 'DNS in CMDB + '])

workbook.oci['IP in CMDB Count']=''
workbook.oci['IP not in CMDB Count']=''
workbook.oci['DNS not in CMDB Count']= ''
workbook.oci['DNS in CMDB Count']=''
workbook.oci['DNS in CMDB + Count']=''

workbook.oci['IP in CMDB Count'][0]= len(workbook.oci[workbook.oci['IP in CMDB']== 'IP in CMDB'])
workbook.oci['IP not in CMDB Count'][0]= len(workbook.oci[workbook.oci['IP in CMDB']== 'IP NOT in CMDB'])
workbook.oci['DNS in CMDB Count'][0]= len(workbook.oci[workbook.oci['DNS in CMDB']== 'DNS in CMDB'])
workbook.oci['DNS not in CMDB Count'][0]= len(workbook.oci[workbook.oci['DNS in CMDB']== 'DNS NOT in CMDB'])
workbook.oci['DNS in CMDB + Count'][0]= len(workbook.oci[workbook.oci['DNS in CMDB']== 'DNS in CMDB + '])

workbook.openair['IP in CMDB Count']=''
workbook.openair['IP not in CMDB Count']=''
workbook.openair['DNS not in CMDB Count']= ''
workbook.openair['DNS in CMDB Count']=''
workbook.openair['DNS in CMDB + Count']=''

workbook.openair['IP in CMDB Count'][0]= len(workbook.openair[workbook.openair['IP in CMDB']== 'IP in CMDB'])
workbook.openair['IP not in CMDB Count'][0]= len(workbook.openair[workbook.openair['IP in CMDB']== 'IP NOT in CMDB'])
workbook.openair['DNS in CMDB Count'][0]= len(workbook.openair[workbook.openair['DNS in CMDB']== 'DNS in CMDB'])
workbook.openair['DNS not in CMDB Count'][0]= len(workbook.openair[workbook.openair['DNS in CMDB']== 'DNS NOT in CMDB'])
workbook.openair['DNS in CMDB + Count'][0]= len(workbook.openair[workbook.openair['DNS in CMDB']== 'DNS in CMDB + '])

workbook.payroll['IP in CMDB Count']=''
workbook.payroll['IP not in CMDB Count']=''
workbook.payroll['DNS not in CMDB Count']= ''
workbook.payroll['DNS in CMDB Count']=''
workbook.payroll['DNS in CMDB + Count']=''

workbook.payroll['IP in CMDB Count'][0]= len(workbook.payroll[workbook.payroll['IP in CMDB']== 'IP in CMDB'])
workbook.payroll['IP not in CMDB Count'][0]= len(workbook.payroll[workbook.payroll['IP in CMDB']== 'IP NOT in CMDB'])
workbook.payroll['DNS in CMDB Count'][0]= len(workbook.payroll[workbook.payroll['DNS in CMDB']== 'DNS in CMDB'])
workbook.payroll['DNS not in CMDB Count'][0]= len(workbook.payroll[workbook.payroll['DNS in CMDB']== 'DNS NOT in CMDB'])
workbook.payroll['DNS in CMDB + Count'][0]= len(workbook.payroll[workbook.payroll['DNS in CMDB']== 'DNS in CMDB + '])

workbook.vpn['IP in CMDB Count']=''
workbook.vpn['IP not in CMDB Count']=''
workbook.vpn['DNS not in CMDB Count']= ''
workbook.vpn['DNS in CMDB Count']=''
workbook.vpn['DNS in CMDB + Count']=''

workbook.vpn['IP in CMDB Count'][0]= len(workbook.vpn[workbook.vpn['IP in CMDB']== 'IP in CMDB'])
workbook.vpn['IP not in CMDB Count'][0]= len(workbook.vpn[workbook.vpn['IP in CMDB']== 'IP NOT in CMDB'])
workbook.vpn['DNS in CMDB Count'][0]= len(workbook.vpn[workbook.vpn['DNS in CMDB']== 'DNS in CMDB'])
workbook.vpn['DNS not in CMDB Count'][0]= len(workbook.vpn[workbook.vpn['DNS in CMDB']== 'DNS NOT in CMDB'])
workbook.vpn['DNS in CMDB + Count'][0]= len(workbook.vpn[workbook.vpn['DNS in CMDB']== 'DNS in CMDB + '])

workbook.blank['IP in CMDB Count']=''
workbook.blank['IP not in CMDB Count']=''
workbook.blank['DNS not in CMDB Count']= ''
workbook.blank['DNS in CMDB Count']=''
workbook.blank['DNS in CMDB + Count']=''

workbook.blank['IP in CMDB Count'][0]= len(workbook.blank[workbook.blank['IP in CMDB']== 'IP in CMDB'])
workbook.blank['IP not in CMDB Count'][0]= len(workbook.blank[workbook.blank['IP in CMDB']== 'IP NOT in CMDB'])
workbook.blank['DNS in CMDB Count'][0]= len(workbook.blank[workbook.blank['DNS in CMDB']== 'DNS in CMDB'])
workbook.blank['DNS not in CMDB Count'][0]= len(workbook.blank[workbook.blank['DNS in CMDB']== 'DNS NOT in CMDB'])
workbook.blank['DNS in CMDB + Count'][0]= len(workbook.blank[workbook.blank['DNS in CMDB']== 'DNS in CMDB + '])

workbook.console['IP in CMDB Count']=''
workbook.console['IP not in CMDB Count']=''
workbook.console['DNS not in CMDB Count']= ''
workbook.console['DNS in CMDB Count']=''
workbook.console['DNS in CMDB + Count']=''

workbook.console['IP in CMDB Count'][0]= len(workbook.console[workbook.console['IP in CMDB']== 'IP in CMDB'])
workbook.console['IP not in CMDB Count'][0]= len(workbook.console[workbook.console['IP in CMDB']== 'IP NOT in CMDB'])
workbook.console['DNS in CMDB Count'][0]= len(workbook.console[workbook.console['DNS in CMDB']== 'DNS in CMDB'])
workbook.console['DNS not in CMDB Count'][0]= len(workbook.console[workbook.console['DNS in CMDB']== 'DNS NOT in CMDB'])
workbook.console['DNS in CMDB + Count'][0]= len(workbook.console[workbook.console['DNS in CMDB']== 'DNS in CMDB + '])

workbook.interface['IP in CMDB Count']=''
workbook.interface['IP not in CMDB Count']=''
workbook.interface['DNS not in CMDB Count']= ''
workbook.interface['DNS in CMDB Count']=''
workbook.interface['DNS in CMDB + Count']=''

workbook.interface['IP in CMDB Count'][0]= len(workbook.interface[workbook.interface['IP in CMDB']== 'IP in CMDB'])
workbook.interface['IP not in CMDB Count'][0]= len(workbook.interface[workbook.interface['IP in CMDB']== 'IP NOT in CMDB'])
workbook.interface['DNS in CMDB Count'][0]= len(workbook.interface[workbook.interface['DNS in CMDB']== 'DNS in CMDB'])
workbook.interface['DNS not in CMDB Count'][0]= len(workbook.interface[workbook.interface['DNS in CMDB']== 'DNS NOT in CMDB'])
workbook.interface['DNS in CMDB + Count'][0]= len(workbook.interface[workbook.interface['DNS in CMDB']== 'DNS in CMDB + '])

workbook.nouse['IP in CMDB Count']=''
workbook.nouse['IP not in CMDB Count']=''
workbook.nouse['DNS not in CMDB Count']= ''
workbook.nouse['DNS in CMDB Count']=''
workbook.nouse['DNS in CMDB + Count']=''

workbook.nouse['IP in CMDB Count'][0]= len(workbook.nouse[workbook.nouse['IP in CMDB']== 'IP in CMDB'])
workbook.nouse['IP not in CMDB Count'][0]= len(workbook.nouse[workbook.nouse['IP in CMDB']== 'IP NOT in CMDB'])
workbook.nouse['DNS in CMDB Count'][0]= len(workbook.nouse[workbook.nouse['DNS in CMDB']== 'DNS in CMDB'])
workbook.nouse['DNS not in CMDB Count'][0]= len(workbook.nouse[workbook.nouse['DNS in CMDB']== 'DNS NOT in CMDB'])
workbook.nouse['DNS in CMDB + Count'][0]= len(workbook.nouse[workbook.nouse['DNS in CMDB']== 'DNS in CMDB + '])

workbook.vm['IP in CMDB Count']=''
workbook.vm['IP not in CMDB Count']=''
workbook.vm['DNS not in CMDB Count']= ''
workbook.vm['DNS in CMDB Count']=''
workbook.vm['DNS in CMDB + Count']=''

workbook.vm['IP in CMDB Count'][0]= len(workbook.vm[workbook.vm['IP in CMDB']== 'IP in CMDB'])
workbook.vm['IP not in CMDB Count'][0]= len(workbook.vm[workbook.vm['IP in CMDB']== 'IP NOT in CMDB'])
workbook.vm['DNS in CMDB Count'][0]= len(workbook.vm[workbook.vm['DNS in CMDB']== 'DNS in CMDB'])
workbook.vm['DNS not in CMDB Count'][0]= len(workbook.vm[workbook.vm['DNS in CMDB']== 'DNS NOT in CMDB'])
workbook.vm['DNS in CMDB + Count'][0]= len(workbook.vm[workbook.vm['DNS in CMDB']== 'DNS in CMDB + '])

workbook.available['IP in CMDB Count']=''
workbook.available['IP not in CMDB Count']=''
workbook.available['DNS not in CMDB Count']= ''
workbook.available['DNS in CMDB Count']=''
workbook.available['DNS in CMDB + Count']=''

workbook.available['IP in CMDB Count'][0]= len(workbook.available[workbook.available['IP in CMDB']== 'IP in CMDB'])
workbook.available['IP not in CMDB Count'][0]= len(workbook.available[workbook.available['IP in CMDB']== 'IP NOT in CMDB'])
workbook.available['DNS in CMDB Count'][0]= len(workbook.available[workbook.available['DNS in CMDB']== 'DNS in CMDB'])
workbook.available['DNS not in CMDB Count'][0]= len(workbook.available[workbook.available['DNS in CMDB']== 'DNS NOT in CMDB'])
workbook.available['DNS in CMDB + Count'][0]= len(workbook.available[workbook.available['DNS in CMDB']== 'DNS in CMDB + '])


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
    
stop = time.perf_counter()

print("It took "+str(round((stop-start)/60,2))+" minutes for the script to complete")