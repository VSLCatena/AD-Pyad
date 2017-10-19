#!/usr/bin/env python

from pyad import *
#from unidecode import unidecode
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
import sys,getopt #https://www.tutorialspoint.com/python/python_command_line_arguments.htm
#import reportlab http://www.reportlab.com/docs/reportlab-userguide.pdf


config = ConfigParser.SafeConfigParser()
config.read('config.cfg')

filename = config.get('general','filename')

server_name = config.get('mail','server_name')
server_port = int(config.get('mail','server_port'))
mail_from = config.get('mail','mail_from')
mail_to = config.get('mail','mail_to')

ou_base = config.get('ou', 'base')
ou_disabled = config.get('ou','disabled')
ou_reunisten = config.get('ou','reunisten')

user_pw=config.get('user','pw')
user_domain=config.get('user','domain')
#template_user_raw=config.get('user','template_user')
if 'template_user_raw' in locals():
 z=template_user_raw.split(',')
 template_user=[]
 for m in z:
  template_user.append(m)

script_dir = os.path.dirname(__file__)
stats=[[[],[]],[[],[]],[[],[]],[[],[]],[]] #create,edit,move with each pos and neg result, add numeric stats at end
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
 "report":1
 }
 try:
  opts, args = getopt.getopt(argv,"hm:d:r:f",["mode=","debug=","report="])
 except getopt.GetoptError:
  print 'Invalid argument'
  print 'Type -h for help'
  sys.exit(2)
 for opt, arg in opts:
  if opt == '-h':
   print 'USAGE:\n'
   print 'Python_script_AD.py -m [cema] -d [0,1] -r [0,1,2] -f'
   print '(m)ode: (c)reate,(e)dit,(m)ove,(a)ctivate'
   print '(d)ebug: 0=console+log 1=log'
   print '(r)eport: 0=none, 1=local, 2=local+mail '
   print '(f)orce ignore warning'
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

def username(name):
# intermezzo create username
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

 if (la_users[0]!=['description','name','phone','mail']):
  str_header=(', '.join('"' + item + '"' for item in la_users[0]))
  log(timestamp("CSV-header incorrect. Header is ["+str_header+"]. Exiting"))
  quit()
 del la_users[0] #remove header
 log(timestamp("CSV-header is correct. LA-users have been imported into memory"))
 if 'template_user' in locals():
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
 log(timestamp("AD-users have been imported"))
 return ad_users



#
# Actions
# 
def create(desc,name,phone=None,mail=None,en=False,pw=user_pw,domain=user_domain):
 try:
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
  stats[0][0].append(name)
 except Exception as e:
  log(timestamp("Error in function (create) for ("+name+"): "+str(e)))
  stats[0][1].append(name+" | Error: "+str(e)) 
 return

def update(desc,name,phone,mail):
 "update attributes (if user in AD and in LA)"
 try: 
  user = pyad.aduser.ADUser.from_cn(name)
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
  log(timestamp("Error in function (update) for ("+name+"): "+str(e)))
  stats[1][1].append(name+" | Error: "+str(e))
  
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
 
 
 
 try:
  
  user = pyad.aduser.ADUser.from_cn(user_cn,ou_base)
  if (enabled == 514) or (enabled == 66050):
   ou = pyad.adcontainer.ADContainer.from_dn(ou_disabled)
   user.move(ou)
  
  if (enabled == 512) or (enabled == 66048):
   ou = pyad.adcontainer.ADContainer.from_dn(ou_reunisten)
   user.disable()
   user.move(ou)   
  stats[2][0].append(user_cn)     
 except Exception as e:
  if (str(e[2][2])=="There is no such object on the server.\r\n"):
   stats[2][0].append(user_cn)
   log(timestamp("Error in function (move) for ("+user_cn+"): "+str(e)))    
  else:
   log(timestamp("Error in function (move) for ("+user_cn+"): "+str(e)))
   stats[2][1].append(user_cn+" | Error: "+str(e))
 return

def wake_up(user_cn):
 try: 
  user = pyad.aduser.ADUser.from_cn(user_cn,ou_base)
  user.set_password(user_pw)
  user.enable()
  stats[3][0].append(user_cn)
 except Exception as e:
  log(timestamp("Error in function (wake_up) for ("+user_cn+"): "+str(e)))
  stats[3][1].append(user_cn+" | Error: "+str(e))
 
 


def filter_list(dataset,set,pos):
#filters the dataset with set which corresponds with position.
#returns new list with filtered results
 list_set=list(set)
 filtered_list =[]
 for data in dataset:
  for item in set:
   if item==data[pos]:
	filtered_list.append(data)

 return filtered_list

def clean(user_raw):
 #remove email and phone number from report, cleanup for report.
 #[u'20-115', u'Bas Aapje', u'012345', u'bas@aapje.local', u'baapje', 512]
 # 20-115 | Bas Aapje
 #["52-000","_TEMPLATE_ Catena Lid","",""]
 #[ | " | (0)
 cleaned = str(user_raw[0])+" | "+str(user_raw[1])
 try:
  cleaned = cleaned+ " | ("+str(user_raw[5])+")"
 except:
  "1"
 return cleaned 
 
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
  f.write(stat+"\n")
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
 f.write("Moved users: "+str(len(stats[2][0]))+" | Errors: "+str(len(stats[2][1]))+"\n")
 f.write("Disabled users: "+str(len(stats[3][0]))+" | Errors: "+str(len(stats[3][1]))+"\n")
 f.write("\n== Details ==\n")
 
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
 f.write("== Enable/Succesfull ==\n")
 for user_ms in stats[3][0]:
  f.write(str(user_ms)+"\n")
 f.write("\n") 
 f.write("== Enable/Unsuccesfull ==\n")
 for user_mu in stats[3][1]:
  f.write(str(user_mu)+"\n")
 f.write("\n") 
 f.close()
 log(timestamp("Report has been written and saved to file"))
 return report_filename
 

def send_mail(filename=None):
 server = smtplib.SMTP(server_name,server_port,'catena-194-171-106-2.leidenuniv.nl')
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
 response=""
 if options["force"]!=True:
  while (response!= "Yes"):
   response=raw_input("Continue? [Yes,No] ")
   if (response=="No"):
    quit()   
 
def execution():
 warning()

 if options["create"]==True:
  for user in list_create:
   create(user[0],user[1],user[2],user[3],True)


  
 if options["edit"]==True:
  for user in list_edit:
   update(user[0],user[1],user[2],user[3])

 
 if options["move"]==True:
  for user in list_move:
   move(user[1],user[5])
 
 if options["enable"]==True:
  for user in list_disabled:
   wake_up(user[1])
 
def post():
 log(timestamp("Created users: "+str(len(stats[0][0]))+" | Errors: "+str(len(stats[0][1]))))
 log(timestamp("Edited users: "+str(len(stats[1][0]))+" | Errors: "+str(len(stats[1][1]))))
 log(timestamp("Moved users: "+str(len(stats[2][0]))+" | Errors: "+str(len(stats[2][1]))))
 log(timestamp("Disabled users: "+str(len(stats[3][0]))+" | Errors: "+str(len(stats[3][1]))))
 if options["report"]>=1:
  report_filename=write_report()
  if options["report"]==2:
   send_mail(report_filename)
  
   
def pre(): 
 la_users=readcsv(filename) 
 ad_users=readAD()
 
#create set from dataset
 s=set(column(la_users,0)) #membership number
 t=set(column(ad_users,0)) #membership number. 


 #do set functions
 c = s - t #create based on LA - new member
 m = t - s #move based on AD - old member
 e = s & t #edit based on LA - current member
 list_create=filter_list(la_users,c,0) #list of strings???
 list_edit=filter_list(la_users,e,0) 
 list_move=filter_list(ad_users,m,0)
 list_disabled=[x for x in ad_users if x[5] in [514,66050]]
 list_enabled=[x for x in ad_users if x[5] in [512,544,66048]];
 #display stats
 stats[4].append("Users in LA: "+str(len(s)))
 stats[4].append("Users in AD: "+str(len(t)))
 stats[4].append("Creatable: "+str(len(c)))
 stats[4].append("Editable: "+str(len(e)))
 stats[4].append("Moveable: "+str(len(m)))
 stats[4].append("Disabled users: "+str(len(list_disabled))+" | Enabled users: "+str(len(list_enabled)))
 for x in stats[4]:
  log(timestamp(x))
 return (list_create,list_edit,list_move,list_disabled, list_enabled)  

#Magic
list_create,list_edit,list_move,list_disabled, list_enabled= pre()
execution()
post()
