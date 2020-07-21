import paramiko
import getpass
import datetime

username = "admin"
password = "admin123"
port = 22

ip_file = open("ipler.txt","r")
ips= ip_file.readlines()
ip_list = []

today = datetime.date.today()

for i in ips:
	ip_list.append(i.strip())
ip_file.close()

#SSH connection parameters
ssh = paramiko.SSHClient()
ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())		

#exec command
for ip in ip_list:
	print("*"*30,ip,": Komut gönderiliyor:")
	ssh.connect(ip,port,username,password,timeout=10)
	stdin,stdout,stderr = ssh.exec_command("sh run")
	result = stdout.read()
	output_file = open(ip+"-"+str(today)+"-backup.txt","a")
	output_file.write(result.decode("utf-8")+"\n\n")
		
output_file.close()
ssh.close()
