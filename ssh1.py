import paramiko
from openpyxl import Workbook,load_workbook
import sys, os, re
import subprocess
import getpass

ip_list=[]
device_name=[]
active_devices=[]
passive_devices=[]
ssh_successful_devices=[]
ssh_failed_for_authentication_devices=[]
ssh_failed_devices=[]
book="test.xlsx"
log = open("log.txt","a",encoding="utf-8")
wb = load_workbook(book)
ws1 = wb["Sayfa1"]
ws2 = wb["Sayfa2"]

port = 22

#*******************************************excel'den ip adreslerini alma

def create_ip_list():
	try:
		for cell in ws1['D']:
			ip_list.append(str(cell.value))
		for cell in ws1['C']:
			device_name.append(str(cell.value))
		wb.close()
		return 1
	except:
		wb.close()
		return 0

#******************************************aktif pasif cihaz listesi olusturma

def active_passive_device_list():

	for ip in range(1,len(ip_list)):
		connectivity = check_ping(ip_list[ip])
		if(connectivity == 1):
			active_devices.append(ip_list[ip])
		else:
			passive_devices.append(ip_list[ip])

#******************************************erisim kontrolu

def check_ping(ip):
	try:
		response = subprocess.Popen(["ping", "-n", "2", "-w", "200", ip]).wait()
		#print(response)
		if response == 1:
			print ("*"*70 + ip + " - inactive!")
			return 0
		elif response == 0:
			print ("*"*70 + ip + " - active!")
			return 1
	except:
		print("There was a problem with the ping test")
		return 0

#*******************************************************ssh baglantisi controlu

def ssh_connect_status(device_ip):

	while True:
		try:
			ssh = paramiko.SSHClient()
			ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
			ssh.connect(device_ip, port, username, password, look_for_keys=False)
			print("*"*70+"Authentication verified for "+ device_ip)
			ssh_successful_devices.append(device_ip)
			return 1
		except paramiko.AuthenticationException as e:
			print("*"*70+"Authentication failed when connecting to "+ device_ip)
			ssh_failed_for_authentication_devices.append(device_ip)
			return 0
		except:
			print("*"*70+"Could not SSH to "+device_ip)
			ssh_failed_devices.append(device_ip)
			return 2
	ssh.close()
	return

#******************************************************inventory bilgisini alıp excel'e yazma

def get_inv(device_ip, row_number,chassis_sn_counter,power_sn_counter):

	ssh = paramiko.SSHClient()
	ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
	ssh.connect(device_ip, port, username, password, look_for_keys=False)
	stdin,stdout,stderr = ssh.exec_command("show inventory")
	output = stdout.readlines()
	ssh.close()
	row = [[]] 
	for line in output: 
	    tmp = []
	    for element in line[0:-1].split(', '):
	        tmp.append(str(element))
	    row.append(tmp)

	counter=1
	chassis_sn_add = 0
	power_sn_add = 0
	x=row_number
	ws2['A1']="Device IP"
	ws2['B1']="Device Name"
	ws2 ['C1'] = "Name"
	ws2['D1']="PID"
	ws2['E1']="DESCR"
	ws2['F1']="SN"	
	
	while counter < len(row):
			name = str(row[counter][0]).replace("NAME: ","")
			descr = str(row[counter][1]).replace("DESCR: ","")
			
			#************************************************************chassis seri no
			if chassis_sn_add < 4 and "chassis" in name.lower():
				chassis_sn = str(row[counter+1][2]).replace("SN: ","")
				if chassis_sn_add == 0:
					ws1['E'+str(chassis_sn_counter+1)] = str(chassis_sn)
				elif chassis_sn_add == 1:
					ws1['F'+str(chassis_sn_counter+1)] = str(chassis_sn)
				elif chassis_sn_add == 2:
					ws1['G'+str(chassis_sn_counter+1)] = str(chassis_sn)
				elif chassis_sn_add == 3:
					ws1['H'+str(chassis_sn_counter+1)] = str(chassis_sn)
				else: 
					print("Stack device more than 4 chassis, please check this device!!"+str(device_ip)+"!!")
					log.write("Stack device more than 4 chassis, please check this device!!"+str(device_ip)+"!!")
				chassis_sn_add += 1
				print("chassis sn add for "+str(device_ip)+" "+str(x)+" "+str(chassis_sn_counter))

			#***********************************************************power seri no

			elif power_sn_add < 2 and "wan" in descr.lower():
				power_pid = str(row[counter+1][0]).replace("PID: ","")
				power_sn = str(row[counter+1][2]).replace("SN: ","")
				if power_sn_add == 0:
					ws1['I'+str(power_sn_counter+1)] = str(power_pid)
					ws1['J'+str(power_sn_counter+1)] = str(power_sn)
				elif power_sn_add == 1:
					ws1['K'+str(power_sn_counter+1)] = str(power_pid)
					ws1['L'+str(power_sn_counter+1)] = str(power_sn)
				else: 
					print("Stack device more than 2 power, please check this device!!"+str(device_ip)+"!!")
					log.write("Stack device more than 2 power supply, please check this device!!"+str(device_ip)+"!!")
				power_sn_add += 1
				print("*****power sn add for "+str(device_ip)+" "+str(x)+" "+str(power_sn_counter))

			#*********************************************************diğer seri nolar

			else:
				str1 = str(row[counter][0]).replace("NAME: ","")
				ws2['D'+str(x)]=(str1)
				str2 = str(row[counter+1][0]).replace("PID: ","")
				ws2['D'+str(x)]=(str2.replace("\"",""))
				print("PID add for "+str(device_ip)+" "+str(x)+" "+str(counter))
				str3 = str(row[counter][1]).replace("DESCR: ","")
				print("DESCR add for "+str(device_ip)+" "+str(x)+" "+str(counter))
				ws2['E'+str(x)]=str(str3.replace("\"",""))
				str4 = str(row[counter+1][2]).replace("SN: ","")
				ws2['F'+str(x)] = str(str4)
				print("SN add for "+str(device_ip)+" "+str(x)+" "+str(counter))
				x+=1

			counter+=3 # inventory bilgileri arasında boş satır yoksa 2'ye düşür!!! 
			print("power sn counter"+str(power_sn_counter))
	return x

#*******************************************************main process

login_inf_counter = 0
login_inf=0
while login_inf == 0 and login_inf_counter < 3:
	username = input("SSH username: ")
	password = getpass.getpass('SSH Password: ')
	if username =="" or password =="":
		print("Please enter a valid username and password.")
		login_inf_counter+=1
	else:
		login_inf=1
row_number=2
if(login_inf == 1):
	read_excel_control = create_ip_list()
	if(read_excel_control == 1):	
		active_passive_device_list()
		for x in range(0,len(active_devices)):
			result_ssh_connect = ssh_connect_status(active_devices[x])
		for x in range(1,len(ip_list)):
			ws2['A'+str(row_number)]=str(ip_list[x])
			ws2['B'+str(row_number)]=str(device_name[x])
			for y in range(0,len(ssh_successful_devices)):
				if (str(ip_list[x]) == str(ssh_successful_devices[y])):
					row_number = get_inv(ssh_successful_devices[y],row_number,x,x)
					row_number -=1				
			row_number +=1	
			wb.save("inventory.xlsx")	
			print("Inventory information saved to inventory.xlsx !!!!")
		wb.close()		
	else:
		print("Ip list is not defined.")
		
else:
	print("Input error!!")
	sys.exit()