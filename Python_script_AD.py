#!/usr/bin/env python

from pyad import *
from unidecode import unidecode
import ConfigParser
import time
import datetime
import csv
import smtplib
from os.path import basename
from email.MIMEMultipart import MIMEMultipart
from email.MIMEText import MIMEText
from email.mime.application import MIMEApplication
import os
#import sys,getopt 
#https://www.tutorialspoint.com/python/python_command_line_arguments.htm
#import reportlab http://www.reportlab.com/docs/reportlab-userguide.pdf


config = ConfigParser.SafeConfigParser()
config.read('config.cfg')

filename = config.get('general','filename')
debuglevel = config.get ('general','debuglevel')
debuglevel = int(debuglevel)
server_name = config.get('mail','server_name')
server_port = config.get('mail','server_port')

ou_base = config.get('ou', 'base')
ou_disabled = config.get('ou','disabled')
ou_reunisten = config.get('ou','reunisten')
user_pw=config.get('user','pw')
user_domain=config.get('user','domain')
template_user=config.get('user','template_user')

script_dir = os.path.dirname(__file__)
stats=[[[],[]],[[],[]],[[],[]],[]] #create,edit,move with each pos and neg result, add numeric stats at end



def log(data):
 #debuglevel
 # 0 write to file and console
 # 1 write to file
 ts=time.time()
 st = datetime.datetime.fromtimestamp(ts).strftime('%Y-%m-%d')
 log_filename=script_dir+"logs/PyAD_"+st+".log"
 f = open(log_filename, 'a+')
 f.write(data+'\n') 
 f.close()
 if (debuglevel == 0):
  print(data)
  

def timestamp(string=""):
 ts=time.time()
 st = datetime.datetime.fromtimestamp(ts).strftime('%Y-%m-%d %H:%M:%S:%f')+" "
 data = st+string
 return data
 
log(timestamp("Starting.... Settings are loaded"))

def username(name):
# intermezzo create username
# 1. fist letter of whole name
# 2. trim of "-" and first " " and name
# 3. count amount of spaces left (0=simple name. >0=complex)
# 4. if complex, count length of second word 
# 5. if longer than 3, remove it
# 6. remove other spaces and add first letter
	#adriaan bart-jan van der hoven schouwen
    name=name.replace("-", "") #Adriaan BartJan van der Hoven Schouwen
    initial=name[0] #A
    pos_space=name.find(" ") #7
    temp_name=name[pos_space+1:] #BartJan van der Hoven Schouwen
	
    count_space=temp_name.count(' ') #4
    if (count_space>0):
     len_2word=len(temp_name[:temp_name.find(" ")]) #7
     if (len_2word > 3):
      temp_name=temp_name[temp_name.find(" ")+1:]#van der Hoven Schouwen
	 
    temp_name = temp_name.replace(" ", "") #vanderHovenSchouwen
    username_raw=initial+temp_name # AvanderHovenSchouwen
    username=username_raw.lower()  #avanderhovenschouwen
    return username
	

def column(matrix, i):
    return [row[i] for row in matrix]

def readcsv(filename):
#
#  read CSV
# header: lidnummer	naam	telefoonnummer	emailadres
# ['descr','name',phone','email']	new
#
#
 la_users=[]

 with open(filename, 'rb') as f:
  reader = csv.reader(f,delimiter = ';')
  mycsv = list(reader)
  for users in mycsv:
   la_users.append(users)

 if (la_users[0]!=['lidnummer','naam','telefoonnummer','emailadres']):
  str_header=(', '.join('"' + item + '"' for item in la_users[0]))
  log(timestamp("CSV-header incorrect. Header is ["+str_header+"]. Exiting"))
  quit()
 del la_users[0] #remove header
 log(timestamp("CSV-header is correct"))
 la_users.append(template_user)
 return la_users


 

def readAD():
 #
 #read AD
 #
 #
 ad_users = []
 import pyad.adquery
 q = pyad.adquery.ADQuery()
 q.execute_query(
     attributes = [ "description", "displayName", "mail", "sAMAccountName","userAccountControl","telephoneNumber"],
     where_clause = "objectClass = 'user'",
     base_dn = ou_base
 )

 for x in q.get_results():
  ad_user_data= []
  tuple_descr=(x["description"]) #Descr is tuple
  ad_user_data.append(tuple_descr[0]) #description is list
  ad_user_data.append(x["displayName"])
  ad_user_data.append(x["telephoneNumber"])
  ad_user_data.append(x["mail"])
  ad_user_data.append(x["sAMAccountName"])
  ad_user_data.append(x["userAccountControl"])
  #try to use GUID
  ad_users.append(ad_user_data)
 return ad_users



#
# Actions
# 
def create(desc,name,phone=None,mail=None,en=False,pw=user_pw,domain=user_domain):
 un=username(name)
 ou_test = pyad.adcontainer.ADContainer.from_dn(ou_base)
 new_user=pyad.aduser.ADUser.create(name, ou_test, password=pw, enable=en)
 new_user.update_attribute("displayName",name) #same as above
 new_user.update_attribute("mail",mail)
 new_user.update_attribute("telephoneNumber",phone)
 new_user.update_attribute("userPrincipalName",un+"@"+domain)
 new_user.update_attribute("sAMAccountName",un)
 new_user.update_attribute("description",desc)
 new_user.update_attribute("profilePath","\\\\"+domain+"\\dfs\\Profiles\\"+un)
 new_user.set_user_account_control_setting('PASSWD_NOTREQD',False)
 stats[0][0].append(user)
 return

def update(desc,name,phone,mail):
 "update attributes (if user in AD and in LA)"
 try: 
  user = pyad.aduser.ADUser.from_cn(names)
  old_mail=user.get_attribute("mail")[0]
  old_phone=user.get_attribute("telephoneNumber")[0]
  #user.clear_attribute("mail")
  #user.append_to_attribute("mail", mail)
  user.update_attribute("description",desc)
  updated_mail = False
  updated_phone = False
  if (mail!=old_mail):
   user.update_attribute("mail", mail)
   updated_mail = True
  if (phone!=old_phone):
   user.update_attribute("telephoneNumber",phone)
   updated_phone = True
  update = "Updated mail: "+str(updated_mail)+" | Updated phone: "+str(updated_phone)
  
  info=[desc,name,update]
  stats[1][0].append(info)
 except Exception as e:
  print(e)
  "1"; 
  stats[1][1].append(user)
  
 return 
   
def move(user,enabled):
 "Move (if user in AD and not in LA) to reunisten if user has been enabled before else to lid af"
 
 # Acc enabled or not (userAccountControl)
 # 512=Enabled
 # 514= Disabled
 # 544= Enabled without password http://activedirectoryfaq.com/2013/12/empty-password-in-active-directory-despite-activated-password-policy/
 #https://msdn.microsoft.com/en-us/library/aa772300.aspx
 
 # 66048 = Enabled, password never expires
 # 66050 = Disabled, password never expires
 
 user = pyad.aduser.ADUser.from_cn(user,ou_base)
 
 try:
  if (enabled == 514) or (enabled == 66050):
   ou = pyad.adcontainer.ADContainer.from_dn(ou_disabled)
   user.move(ou)
  
  if (enabled == 512) or (enabled == 66048):
   ou = pyad.adcontainer.ADContainer.from_dn(ou_reunisten)
   user.disable()
   user.move(ou)   
  stats[2][0].append(user)     
 except Exception as e:
  print(e)
  stats[2][1].append(user)
 return

def wake_up(user):
 password=user_pw
 user.set_password(password)
 user.enable()
 


def filter_list(dataset,set,pos):
#filters the dataset with set which corresponds with position.
#returns new list with filtered results
 list_set=list(set)
 filtered_list =[]
 for data in dataset:
  for item in set:
   if item==data[pos]:
	filtered_list.append(data)
   item==data[pos]

 return filtered_list

def write_report():
 ts=time.time()
 st = datetime.datetime.fromtimestamp(ts).strftime('%Y-%m-%d %H-%M-%S')
 report_filename=script_dir+"reports/PyAD_Report_"+st+".txt"
 f = open(report_filename,'w+')
 f.write("Report of Active directory modifications\n")
 f.write("Date: "+timestamp()+"\n\n")
 f.write("== Statistics before edit ==\n")
 for stat in stats[3]:
  f.write(stat+"\n")
 f.write("\n")
 f.write("== Users to be created ==\n")
 for user_c in list_create:
  f.write(str(user_c)+"\n")
 f.write("\n")
 f.write("== Users who are editable ==\n")
 for user_e in list_edit:
  f.write(str(user_e)+"\n")
 f.write("\n") 
 f.write("== Users to be deactivated ==\n")
 for user_m in list_move:
  f.write(str(user_m)+"\n")
 f.write("\n")
 f.write("== Enabled users ==\n")
 for user_e in list_enabled:
  f.write(str(user_e)+"\n")
 f.write("\n")
 f.write("== Disabled users ==\n")
 for user_d in list_disabled:
  f.write(str(user_d)+"\n")
 f.write("\n")
 
 f.write("== Statistics after edit ==\n")
 f.write("== Create/Succesfull ==\n")
 for user_cs in stats[0][0]:
  f.write(str(user_cs)+"\n")
 f.write("\n")
 f.write("== Create/Unsuccesfull ==\n")
 for user_cu in stats[0][1]:
  f.write(str(user_cu)+"\n")
 f.write("\n")
 f.write("== Edit/Succesfull ==\n")
 for user_es in stats[1][0]:
  f.write(str(user_es)+"\n")
 f.write("\n") 
 f.write("== Edit/Unsuccesfull ==\n")
 for user_cu in stats[1][1]:
  f.write(str(user_eu)+"\n")
 f.write("\n") 
 f.write("== Move/Succesfull ==\n")
 for user_ms in stats[2][0]:
  f.write(str(user_ms)+"\n")
 f.write("\n") 
 f.write("== Move/Unsuccesfull ==\n")
 for user_mu in stats[2][1]:
  f.write(str(user_mu)+"\n")
 f.write("\n")
 
 f.close()
 
def send_mail(filename=None):
 server = smtplib.SMTP(server_name,server_port)
 msg = MIMEMultipart()
 msg['From'] = mail_from
 msg['To'] = mail_to
 msg['Subject'] = "Active Directory Report ("+timestamp()+")"
 body = "Python test mail"
 msg.attach(MIMEText(body, 'plain'))
 if filename!=None:
  with open(filename, "rb") as fil:
   part = MIMEApplication(fil.read(),Name=basename(filename))
  # After the file is closed
  part['Content-Disposition'] = 'attachment; filename="%s"' % basename(filename)
  msg.attach(part)
 server.ehlo()
 server.starttls()
 server.ehlo()
 text=msg.as_string()
 server.sendmail(mail_from,mail_to,text)

 
la_users=readcsv(filename) 
ad_users=readAD()
 
#create set from dataset
s=set(column(la_users,0)) #membership number
t=set(column(ad_users,0)) #membership number. 


#do set functions
c = s - t #create based on LA - new member
m = t - s #move based on AD - old member
e = s & t #edit based on LA - current member
 
list_create=filter_list(la_users,c,0)
list_edit=filter_list(la_users,e,0) 
list_move=filter_list(ad_users,m,0)
list_disabled=[x for x in ad_users if x[5] in [514,66050]]
list_enabled=[x for x in ad_users if x[5] in [512,544,66048]];


#display stats
stats[3].append("Users in LA: "+str(len(s)))
stats[3].append("Users in AD: "+str(len(t)))
stats[3].append("Creatable: "+str(len(c)))
stats[3].append("Editable: "+str(len(e)))
stats[3].append("Moveable: "+str(len(m)))
stats[3].append("Disabled users: "+str(len(list_disabled))+" | Enabled users: "+str(len(list_enabled)))

for x in stats[3]:
 log(timestamp(x)) 
 
 
#
# Execute functions 
#


raw_input("Continue? ") 
raw_input("Are you sure? ") 
# for user in list_create:
 # try:
  # create(user[0],user[1],user[2],user[3],True)
 # except Exception as e:
  # print e
  #log(timestamp(e))
	
# for user in list_edit:
 # try:
  # update(user[0],user[1],user[2],user[3])
 # except Exception as e:
  # print(e)

# for user in list_move:
 # try:
  # move(user[1],user[5])
 # except Exception as e:
  # print(e)
write_report()
