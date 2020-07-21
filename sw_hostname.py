import paramiko
import time
from openpyxl import *

user = "admin"
password = "admin123"
port = 22

book="sw_hostname.xlsx"
wb = load_workbook(book, data_only=True)
ws1 = wb["SW List"]

commands = ['conf t\n', 'host deneme5\n', 'do wr\n']

ip_list = []
sw_hostname = []

for cell1 in ws1['A']:
	ip_list.append(str(cell1.value))
for cell2 in ws1['B']:
    sw_hostname.append(str(cell2.value))
    

ssh = paramiko.SSHClient()
ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())


for x in range(1, len(ip_list)):
    ip = ip_list[x]
    print("*"*30,ip,": Komut gönderiliyor:")
    ssh.connect(ip,port,user,password,timeout=10)
    channel = ssh.invoke_shell()
    commands[1] = "hostname "+sw_hostname[x]+"\n"
    for i in range(0,len(commands)): 
        channel.send(commands[i])       
        while not channel.recv_ready():
            time.sleep(1)
        out = channel.recv(2048)
    ssh.close()