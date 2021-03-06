import paramiko
import sys, os, re
from openpyxl import *
import subprocess
import getpass
import datetime

ip_list=[]
active_devices=[]
passive_devices=[]
ssh_successful_devices=[]
ssh_failed_for_authentication_devices=[]
ssh_failed_devices=[]
port = 22


#*********************************************************excel listelerini tanımlama ve düzenleme

book="test.xlsx"
wb = load_workbook(book, data_only=True)
ws = wb["sw list"]

#********************************************************excel'den ip adreslerini alma

def create_ip_list():
	try:
		for cell in ws['P']:
			ip_list.append(str(cell.value))
		wb.close()
		return 1
	except:
		wb.close()
		return 0

#********************************************************aktif pasif cihaz listesi olusturma

def active_passive_device_list():

	for ip in range(1,len(ip_list)):
		connectivity = check_ping(ip_list[ip])		
		if(connectivity == 1):
			active_devices.append(ip_list[ip])
		else:
			passive_devices.append(ip_list[ip])
	return

#*******************************************************erisim kontrolu

def check_ping(ip):
	try:
		response = subprocess.Popen(["ping.exe", ip], stdout = subprocess.PIPE).communicate()[0]
		#print(response)
		if ("unreachable" in str(response)):
			print (ip + "  is unreachable!")
			return 0
		elif ("timed out." in str(response)):
			print (ip + "  is unreachable!")
			return 0
		elif response == 1:
			print (ip + " is unreachable!")
			return 0
		else:
			print (ip + " is reachable!")
			return 1
	except:
		print("There was a problem with the ping test")
		return 0

#*******************************************************ssh baglantisi kontrolu

def ssh_connect_status(device_ip):

	while True:
		try:
			ssh = paramiko.SSHClient()
			ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
			ssh.connect(device_ip, port, username, password, look_for_keys=False)
			print("*"*70+"Authentication verified for "+ device_ip)
			ssh_successful_devices.append(device_ip)
			return 1
		except paramiko.AuthenticationException:
			print("*"*70+"Authentication failed when connecting to "+ device_ip)
			ssh_failed_for_authentication_devices.append(device_ip)
			return 0
		except paramiko.SSHException:
			print("*"*70+"Unable to establish SSH connection: " + device_ip)
			ssh_failed_devices.append(device_ip)
			return 0
		except:
			print("*"*70+"Unable to establish SSH connection: " + device_ip)
			ssh_failed_devices.append(device_ip)
			return 0
	ssh.close()
	return

#******************************************************backup dosyası oluşturma

def get_backup(device_ip):

	now = datetime.datetime.now()
	ssh = paramiko.SSHClient()
	ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
	ssh.connect(device_ip, port, username, password, look_for_keys=False)
	stdin,stdout,stderr = ssh.exec_command("show running-config")
	output = stdout.read()
	ssh.close()
	os.system("mkdir backup")
	path = "./backup/"+str(device_ip)+"-"+str(now.day)+str(now.month)+str(now.year)+"-"+"backup.txt"
	mode = 0o666
	flags = os.O_RDWR | os.O_CREAT 
	fd = os.open(path, flags, mode) 
	os.write(fd, output) 
	os.close(fd)
	return

#*****************************************************işlem yapılamayan cihazlar
def print_failed_devices(list_name, xlist):

		path = "./backup/"+"backup_failed_devices.txt"
		mode = 0o666
		flags = os.O_RDWR | os.O_CREAT 
		fd = os.open(path, flags, mode)
		print(list_name)
		name = str.encode(list_name+"\n")
		os.write(fd, name)
		for x in range(0, len(xlist)):
			print(xlist[x])
			ip = str.encode(xlist[x]+"\n")
			os.write(fd, ip) 
		os.close(fd)						
		return

#******************************************************main process
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
		if active_devices:
			for x in range(0,len(active_devices)):
				result_ssh_connect = ssh_connect_status(active_devices[x])
			for x in range(1,len(ip_list)):
				for y in range(0,len(ssh_successful_devices)):
					if (str(ip_list[x]) == str(ssh_successful_devices[y])):
						get_backup(ip_list[x])				
		if passive_devices:	
			print_failed_devices("passive_devices", passive_devices)
		if ssh_failed_for_authentication_devices:	
			print_failed_devices("ssh_failed_for_authentication_devices",ssh_failed_for_authentication_devices)
		if ssh_failed_devices:	
			print_failed_devices("ssh_failed_devices",ssh_failed_devices)
		print("Inventory information saved !!")
	else:
		print("Ip list is not defined.")
		sys.exit()
		
else:
	print("Input error!!")
	sys.exit()