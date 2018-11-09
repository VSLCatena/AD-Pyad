#!/usr/bin/env python

from pyad import *
from unidecode import unidecode
import ConfigParser
import time
import datetime
import string
import random
import csv
import smtplib
import pyad.adquery
import logging
from os import rename
from os.path import basename
from email.MIMEMultipart import MIMEMultipart
from email.MIMEText import MIMEText
from email.mime.application import MIMEApplication
import os
import sys,getopt #https://www.tutorialspoint.com/python/python_command_line_arguments.htm
#import reportlab http://www.reportlab.com/docs/reportlab-userguide.pdf
from pywintypes import com_error #needed for exceptions.

config = ConfigParser.SafeConfigParser()
config.read('config.cfg')

filename = config.get('general','filename')
special_folder =config.get('general','special_folder')
folder =config.get('general','folder')
csv_del =config.get('general','delimiter')

server_name = config.get('mail','server_name')
server_port = int(config.get('mail','server_port'))
mail_from = config.get('mail','mail_from')
mail_to = config.get('mail','mail_to')
local_hostname=config.get('mail','local_hostname')

ou_base = config.get('ou', 'base')
ou_base_group = config.get('ou','base_group')
ou_disabled = config.get('ou','disabled')
ou_reunisten = config.get('ou','reunisten')

#user_pw=config.get('user','pw')
user_domain=config.get('user','domain')
template_user_raw=config.get('user','template_user')
if 'template_user_raw' in locals() or 'template_user_raw' in globals():
 z=template_user_raw.split(',')
 template_user=[]
 for m in z:
  template_user.append(m)
 template_user[2]=None
 template_user[3]=None

 
#script_dir = os.path.dirname(__file__)+"/"
script_dir = "C:/scripts/Computeraccounts/"
stats=[[[],[]],[[],[]],[[],[]],[[],[]],[[0,0,0]]] #create,edit,move with each pos and neg result, add numeric stats at end
#pyad.set_defaults(ldap_server="dc1.domain.com", username="service_account", password="mypassword")

def arguments(argv):
 error = False
 options={
 "create":False,
 "edit":False,
 "move":False,
 "enable":False,
 "debug":1,
 "force":False,
 "report":1,
 "auto":0
 }
 try:
  opts, args = getopt.getopt(argv,"hm:d:r:fb",["mode=","debug=","report="])
 except getopt.GetoptError:
  print 'Invalid argument'
  print 'Type -h for help'
  sys.exit(2)
 for opt, arg in opts:
  if opt == '-h':
   print 'USAGE:\n'
   print 'Python_script_AD.py -m [cema] -d [0,1] -r [0,1,2] -f -b'
   print '(m)ode: (c)reate,(e)dit,(m)ove,(a)ctivate'
   print '(d)ebug: 0=console+log 1=log'
   print '(r)eport: 0=none, 1=local, 2=local+mail '
   print '(f)orce ignore warning'
   print 'Automate via special folder (b)'
   print 'Default:Python_script_AD.py -m cem -d 0 -r 1' 
   sys.exit()
  elif opt in ("-m", "--mode"):
   for l in arg:
    if (l=="c"):
     options["create"]=True
    elif (l=="e"):
     options["edit"]=True
    elif (l=="m"):
     options["move"]=True
    elif (l=="a"):
     options["enable"]=True	
    if l not in ["c","e","m","a"]:
     error = True	
  elif opt in ("-d", "--debug"):
   if (int(arg) in [0,1]) :
    options["debug"]=int(arg)
   else:
    error=True
  elif opt in ("-r", "--report"):
   if (int(arg) in [0,1,2]):
    options["report"]=int(arg)
   else:
    error=True
  if opt == '-f':
   options["force"]=True
  if opt == '-b':
   options["auto"]=True
 count_args=0 #just to make sure that m,d,r are called  
 for x in opts:  
  if x[0] in ["-d","-r"]:
   count_args+=1  
 
 if (not opts)or count_args!=2 or error:  
  print 'Missing or incorrect argument'
  print 'Type -h for help'
  sys.exit(2)
  
 return options  


if __name__ == "__main__":
   options = arguments(sys.argv[1:])

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
 if (options['debug'] == 0):
  print(data)
  

def timestamp(string=""):
 ts=time.time()
 st = datetime.datetime.fromtimestamp(ts).strftime('%Y-%m-%d %H:%M:%S:%f')+" "
 data = st+str(string)
 return data
 
log(timestamp("Starting.... Settings are loaded"))
log(timestamp("Please check the delimiter when you encounter error"))


def username(name):
# intermezzo create username
# The sAMAccountName attribute must be less than 20 characters
# 1. fist letter of whole name
# 2. trim of "-" and first " " and name
# 3. count amount of spaces left (0=simple name. >0=complex)
# 4. if complex, count length of second word 
# 5. if longer than 3, remove it
# 6. remove other spaces and add first letter
	#adriaan bart-jan van der hoven schouwen
    name=name.replace("'","") #remove characters like ' 
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
    username=username[0:19] #force <20
    return username

def password(N):
 new=''.join(random.SystemRandom().choice(string.ascii_letters + string.digits + string.punctuation) for _ in range(N))
 return new
 
def column(matrix, i):
    return [row[i] for row in matrix]

def readcsv(filename):
#
#  read CSV
# header: name, duplicate_name, username, duplicate_username, lidnummber,phone,email
#  
#
#
 la_users=[]
 try:
  with open(filename, 'rb') as f:
   reader = csv.reader(f,delimiter = csv_del)
   mycsv = list(reader)
   for users in mycsv:
    if users[2]=="": #USERNAME
     users[2]=None
    if users[5]=="": #PHONE
     users[5]=None
    if users[6]=="": #MAIL
	 users[6]=None
    la_users.append(users) 

   if (la_users[0]!=['name','duplicate_name','username','duplicate_username','lidnummer','phone','email']):
    str_header=(', '.join('"' + item + '"' for item in la_users[0]))
    log(timestamp("CSV-header incorrect. Header is ["+str_header+"]. Exiting"))
    quit()
   del la_users[0] #remove header
   log(timestamp("CSV-header is correct. LA-users have been imported into memory"))
   if 'template_user' in globals():
    la_users.append(template_user)
 except IOError as e:
  log(timestamp("No file found. Exiting. "+str(e)+""))
  quit()
 return la_users


 

def readAD():
 #
 #read AD
 #
 #
 ad_users = []

 q = pyad.adquery.ADQuery()
 q.execute_query(
     attributes = [ "description", "displayName", "mail", "sAMAccountName","userAccountControl","telephoneNumber"],
     where_clause = "objectClass = 'user'",
     base_dn = ou_base
 )

 for x in q.get_results():
  ad_user_data= []
  try:
   tuple_descr=(x["description"]) #Descr is tuple
   ad_user_data.append(tuple_descr[0]) #description is list
  except: 
   ad_user_data.append("") #description is empty
  ad_user_data.append(x["displayName"])
  ad_user_data.append(x["telephoneNumber"])
  ad_user_data.append(x["mail"])
  ad_user_data.append(x["sAMAccountName"])
  ad_user_data.append(x["userAccountControl"])
    
  #try to use GUID
  ad_users.append(ad_user_data)
 log(timestamp("AD-users have been imported"))
 return ad_users




def create(desc,name,usr=None,phone=None,mail=None,en=False,pw=password(50),domain=user_domain):
 from pyad.pyadexceptions import win32Exception,comException
 ou_test = pyad.adcontainer.ADContainer.from_dn(ou_base)
 common_group=pyad.adgroup.ADGroup.from_dn(ou_base_group)
 error_info="Error: "
 un_suffix=0
 un_success=False
 if usr==None:
  un=username(name)
 if usr!=None:
  un=usr 
 
 #while loop for CN's versus name from CSV
 
 while (un_success==False): #stop after success or on break
  try:
   new_user=pyad.aduser.ADUser.create(name, ou_test, password=pw, enable=en, optional_attributes={
   "sAMAccountName":un,
   "userPrincipalName":un+"@"+domain,
   "profilePath":"\\\\"+domain+"\\dfs\\Profiles\\"+un,
   "description":desc,
   "displayName":name,
   "mail":mail,
   "telephoneNumber":phone
   })
   un_success=True #un is unique
   new_user.set_user_account_control_setting('PASSWD_NOTREQD',False)
   if new_user.is_member_of(common_group)==False:
    new_user.add_to_group(common_group)
  #2 types of errors are possible: 0x8007001f(too long) 0x80071392(duplicate name/un)
  except win32Exception as e:
   if (str(e) =='0x80071392: The object already exists.\r\n'): #The object already exists un or name.
    if (un_suffix<4):
     un_suffix+=1
     un=username(name)+str(un_suffix)
     error_info="Duplicate username: " #try again
    else:
	 error_info="Duplicate name: "
	 break
 
   if (str(e) =='0x8007001f: A device attached to the system is not functioning.\r\n'): #un too long
    error_info="Username too long: "
    break   
  except Exception as e:
   error_info="Error: "
   break
   
 if (un_success==True):
  info=[desc,name,un]
  stats[0][0].append(info)
 if (un_suffix==4) or (un_success==False):
  info=[desc,name,un,error_info+str(e)]
  stats[0][1].append(info)
  log(timestamp("Error in function (create) for ("+name+"): "+error_info+str(e)))
  
 return

def update(desc,name,phone,mail):
 "update attributes (if user in AD and in LA)"
 try: 
  user = pyad.aduser.ADUser.from_cn(name,ou_base)
  try:
   old_mail=user.get_attribute("mail",False)
   old_phone=user.get_attribute("telephoneNumber",False)
  except:
   old_mail=None
   old_phone=None
  #user.clear_attribute("mail")
  #user.append_to_attribute("mail", mail)
  #user.update_attribute("description",desc) 
  #add to group leden if not already
  try:
   pwlast=user.get_password_last_set()
  except ValueError as e:
   pwlast="Never/Expired"
  
  updated_group = False
  common_group=pyad.adgroup.ADGroup.from_dn(ou_base_group)
  if user.is_member_of(common_group)==False:
   user.add_to_group(common_group)
   updated_group = True
   stats[4][0][0]+=1
  updated_mail = False
  updated_phone = False
  if (mail!=old_mail):
   user.update_attribute("mail", mail)
   updated_mail = True
   stats[4][0][1]+=1
  if (phone!=old_phone):
   user.update_attribute("telephoneNumber",phone)
   updated_phone = True
   stats[4][0][2]+=1
  
  date_create=user.get_attribute("whenCreated",False)
  date_change=user.get_attribute("whenChanged",False)
  dates=[date_create,date_change,pwlast]
  if (updated_phone==True or updated_mail==True or updated_group==True):
   update = [updated_mail,updated_phone,updated_group]
   info=[desc,name]+update+dates
   stats[1][0].append(info)
 except Exception as e:
  info=[desc,name,str(e)]
  log(timestamp("Error in function (update) for ("+name+"): "+str(e)))
  stats[1][1].append(info)
  
 return 
   
def move(user_cn,enabled):

 "Move (if user in AD and not in LA) to reunisten if user has been enabled before else to lid af"
 
 # Acc enabled or not (userAccountControl)
 # 512=Enabled
 # 514= Disabled
 # 544= Enabled without password http://activedirectoryfaq.com/2013/12/empty-password-in-active-directory-despite-activated-password-policy/
 #https://msdn.microsoft.com/en-us/library/aa772300.aspx
 
 # 66048 = Enabled, password never expires
 # 66050 = Disabled, password never expires
 

 info=["",user_cn]
 
 try:
  e=None
  user = pyad.aduser.ADUser.from_cn(user_cn,ou_base)
  try:
   desc=str(user.get_attribute("description",False))
   info[0]=desc
  except:
   pass
  user.set_user_account_control_setting('PASSWD_NOTREQD',False)
  #remove privacy sensitive info
  user.update_attribute("telephoneNumber","")
  user.update_attribute("mail","")
  
  if (enabled in [514,546,66050]):
   ou = pyad.adcontainer.ADContainer.from_dn(ou_disabled)
   user.move(ou)
  if (enabled in [512,544,66048]):
   ou = pyad.adcontainer.ADContainer.from_dn(ou_reunisten)
   user.disable()
   user.move(ou)   
  stats[2][0].append(info)
  
 except com_error as e:
  if str(e[2][2]=="There is no such object on the server.\r\n"):
   #successfull but random error
   stats[2][0].append(info)
  else:
   stats[2][1].append(info+[str(e)])
   
 except StandardError as e:  
  stats[2][1].append(info+[str(e)])
  
 finally:
   log(timestamp("Error in function (move) for ("+user_cn+"): "+str(e)))
 return

def wake_up(user_cn):
 info=["",user_cn]
 try: 
  user = pyad.aduser.ADUser.from_cn(user_cn,ou_base)
  try:
   desc=str(user.get_attribute("description",False))
   info[0]=desc
  except:
   pass
  user.set_user_account_control_setting('PASSWD_NOTREQD',False)
  user.set_password(user_pw)
  user.enable()
  stats[3][0].append(info)
 except Exception as e:
  log(timestamp("Error in function (wake_up) for ("+user_cn+"): "+str(e)))
  stats[3][1].append(info+[str(e)])
 
 


def filter_list(dataset,set,pos):
#filters the dataset with set which corresponds with position.
#returns new list with filtered results
 list_set=list(set)
 filtered_list =[]
 for data in dataset:
  for item in set:
   if (item==data[pos]):
	filtered_list.append(data)
 return filtered_list

def clean(user_raw):
 #remove email and phone number from report, cleanup for report.
 #[u'20-115', u'Bas Aapje', u'012345', u'bas@aapje.local', u'baapje', 512]
 # 20-115 | Bas Aapje
 #["52-000","_TEMPLATE_ Catena Lid","",""]
 #[ | " | (0)
 #
 #	BROKEN DUE TO ACCENT IN NAMES
 #
 cleaned = str(user_raw[1])
 print user_raw
 if (len(user_raw[0])>1):
  cleaned = cleaned+ " | ("+str(user_raw[0])+")"
 try:
  cleaned = cleaned+ " | ["+str(user_raw[5])+"]"
 except:
  pass
 return cleaned 

def write_csv():

 #data is from stats[0:2]
 #
 # 
 # Headers (c=la, e=la, m=ad)
 h = ["Number","Name"]
 hc=h+["Username"]
 he=h+["Updated Mail","Updated Phone","Updated Group","Created","Modified","Pwd changed"]
 hm=h
 header=[hc,he,hm]
 filetype=["Create","Edit","Move"]
 #timestamp
 ts=time.time()
 st = datetime.datetime.fromtimestamp(ts).strftime('%Y-%m-%d')
 for i, group in enumerate(stats[0:3]):
  #for each i create new file.
  file="Reports\\Report "+filetype[i]+" "+st+".csv"
  if options['auto']==True:
   file=special_folder+file
  try:
   with open(file, 'wb') as csvfile:
    f = csv.writer(csvfile, delimiter=';')
    f.writerow(header[i])
    for type in group:
     for user in type:
      f.writerow(user)
  except IOError as e:
   log(timestamp("No file found. Exiting. "+str(e)+""))
   quit()
 return
	
	
	
 
def write_report():
 ts=time.time()
 st = datetime.datetime.fromtimestamp(ts).strftime('%Y-%m-%d %H-%M-%S')
 report_filename=script_dir+"reports/PyAD_Report_"+st+".txt"
 f = open(report_filename,'w+')
 f.write("Report of Active directory modifications\n")
 f.write("Date: "+timestamp()+"\n\n")
 f.write("== Options ==\n")
 for key, value in options.iteritems():
  f.write(str(key)+": "+str(value)+"\n")
 f.write("\n== Statistics before edit ==\n")
 for stat in stats[4]:
  f.write(str(stat)+"\n")
 f.write("\n")
 f.write("== Users to be created (new members) ==\n")
 for user_c in list_create:
  f.write(clean(user_c)+"\n")
 f.write("\n")
 f.write("== Users who are editable (members) ==\n")
 for user_e in list_edit:
  f.write(clean(user_e)+"\n")
 f.write("\n") 
 f.write("== Users to be deactivated (old members) ==\n")
 for user_m in list_move:
  f.write(clean(user_m)+"\n")
 f.write("\n")
 f.write("== Enabled users ==\n")
 for user_e in list_enabled:
  f.write(clean(user_e)+"\n")
 f.write("\n")
 f.write("== Disabled users ==\n")
 for user_d in list_disabled:
  f.write(clean(user_d)+"\n")
 f.write("\n")
 
 f.write("== Statistics after edit ==\n")
 f.write("== Summary ==\n")
 f.write("Created users: "+str(len(stats[0][0]))+" | Errors: "+str(len(stats[0][1]))+"\n")
 f.write("Edited users: "+str(len(stats[1][0]))+" | Errors: "+str(len(stats[1][1]))+"\n")
 f.write("#updated group/mail/phone: "+str(stats[4][0])+"\n")
 f.write("Moved users: "+str(len(stats[2][0]))+" | Errors: "+str(len(stats[2][1]))+"\n")
 f.write("Disabled users: "+str(len(stats[3][0]))+" | Errors: "+str(len(stats[3][1]))+"\n")
 f.write("\nDetailed reports can be viewed in Excel using the CSV-files") 
 f.close()
 log(timestamp("Report has been written and saved to file. Detailed report is being created"))
 write_csv()
 log(timestamp("Detailed report is finished"))
 return report_filename
 

def send_mail(filename=None):
 server = smtplib.SMTP(server_name,server_port,local_hostname)
 msg = MIMEMultipart()
 msg['From'] = mail_from
 msg['To'] = mail_to
 msg['Subject'] = "Active Directory Report ("+timestamp()+")"
 body = "Dear,\n This is your Active Directory Report of your server.\nEnclosed is the report in text file.\nYour administrator"
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
 log(timestamp("E-mail wil be sent"))
 server.sendmail(mail_from,mail_to,text)

def warning():
 print("\nThis will happen if you continue:\n"+str(options)+"\n")
 response=""
 if options["force"]!=True:
  while (response!= "Yes"):
   response=raw_input("Continue? [Yes,No,Details] ")
   if (response=="No"):
    quit()   
   if (response=="Details"):
    response2=raw_input("Which list? [list_create,list_edit,list_move] ")
    if (response2 in ["list_create","list_edit","list_move"]):
     show_list=eval(response2)
     for show in show_list:
	  print show
	   
	
def execution():
 warning()

 if options["create"]==True:
  for user in list_create:
   create(user[4],user[0],user[2],user[5],user[6],True)
   time.sleep(1)


  
 if options["edit"]==True:
  for user in list_edit:
   update(user[4],user[0],user[2],user[5])

 
 if options["move"]==True:
  for user in list_move:
   move(user[1],user[5])
 
 if options["enable"]==True:
  for user in list_disabled:
   wake_up(user[1])
 
def post():
 log(timestamp("Location: "+str(filename)))
 log(timestamp("Created users: "+str(len(stats[0][0]))+" | Errors: "+str(len(stats[0][1]))))
 log(timestamp("Edited users: "+str(len(stats[1][0]))+" | Errors: "+str(len(stats[1][1]))))
 log(timestamp("#updated group/mail/phone: "+str(stats[4][0])))
 log(timestamp("Moved users: "+str(len(stats[2][0]))+" | Errors: "+str(len(stats[2][1]))))
 log(timestamp("Disabled users: "+str(len(stats[3][0]))+" | Errors: "+str(len(stats[3][1]))))
 if options["report"]>=1:
  report_filename=write_report()
  if options["report"]==2:
   send_mail(report_filename)

 rename(filename,filename+"_processed.csv")
 log(timestamp("CSV-file has been renamed"))
  
   
def pre(): 
 la_users=readcsv(filename)
 ad_users=readAD()
 
 
#create set from dataset
 s=set(column(la_users,4)) #membership number
 t=set(column(ad_users,0)) #membership number. 
  
 #do set functions
 c = s - t #create based on LA - new member
 m = t - s #move based on AD - old member
 e = s & t #edit based on LA - current member
 list_create=filter_list(la_users,c,4) #list of strings???
 list_edit=filter_list(la_users,e,4) 
 list_move=filter_list(ad_users,m,0) #problems, m contains one "" and list_move all ""
 list_disabled=[x for x in ad_users if x[5] in [514,546,66050]]
 list_enabled=[x for x in ad_users if x[5] in [512,544,66048]];

#sort by name 
 def getKey(item):
  return item[1] 
 list_move=sorted(list_move, key=getKey) 
 list_disabled=sorted(list_disabled, key=getKey)  
 list_enabled=sorted(list_enabled, key=getKey) 
 
 #debug numbers 
 # f2 = open("numbers.txt",'w+')
 # for nr in list_edit:
  # f2.write(str(nr)+"\n")
 # f2.close()
 
 #display stats
 stats[4].append("Users in LA: "+str(len(s))+" ("+str(len(la_users))+" incl. without nr)")
 stats[4].append("Users in AD: "+str(len(t))+" ("+str(len(ad_users))+" incl. without nr)")
 stats[4].append("Creatable: "+str(len(list_create)))
 stats[4].append("Editable: "+str(len(list_edit)))
 stats[4].append("Moveable: "+str(len(list_move))+" ("+str(len(m)-1)+" users with nr)")
 stats[4].append("Disabled users: "+str(len(list_disabled))+" | Enabled users: "+str(len(list_enabled)))
 for x in stats[4]:
  log(timestamp(x))
 return (list_create,list_edit,list_move,list_disabled, list_enabled)  


if options['auto']==True:
 filename=special_folder+filename
if options['auto']!=True:
 filename=folder+filename  
 
#Magic
list_create,list_edit,list_move,list_disabled, list_enabled= pre()
execution()
post()


