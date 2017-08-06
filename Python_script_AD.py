#!/usr/bin/env python


#
#import
#
from pyad import *
from unidecode import unidecode
import ConfigParser



#
#settings
#
#loading
config = ConfigParser.SafeConfigParser()
config.read('config.cfg')

filename = config.get('general','filename')
ou_base = config.get('ou', 'base')
ou_disabled = config.get('ou','disabled')
ou_reunisten = config.get('ou','reunisten')
user_pw=config.get('user','pw')
user_domain=config.get('user','domain')


#connect to AD
#
#
#  read CSV
# ['name','username','descr',email']
#
import csv
la_users=[]
with open(filename, 'rb') as f:
 reader = csv.reader(f,delimiter = ';')
 #for row in reader:
  #print row[0][0]
 mycsv = list(reader)
 for users in mycsv:
  la_users.append(users)
del la_users[0] #remove header  
la_users_name=[bb[0] for bb in la_users]
  


#
#read AD
#
#
import pyad.adquery
q = pyad.adquery.ADQuery()
q.execute_query(
    attributes = [ "description", "displayName", "mail", "sAMAccountName","userAccountControl"],
    where_clause = "objectClass = 'user'",
    base_dn = ou_base
)
ad_users = []
for x in q.get_results():
 ad_user_data= []
 ad_user_data.append(x["displayName"])
 ad_user_data.append(x["sAMAccountName"])
 ad_user_data.append(x["mail"])
 ad_user_data.append(x["description"])
 ad_user_data.append(x["userAccountControl"])
 ad_users.append(ad_user_data)
ad_users_name=[xx[0] for xx in ad_users]



#
# Actions
# 

def create(name,un,desc=None,mail=None,en=False,pw=user_pw,domain=user_domain ):
 "create (if not in AD and in LA)"
 ou_test = pyad.adcontainer.ADContainer.from_dn(ou_base)
 new_user=pyad.aduser.ADUser.create(name, ou_test, password=pw, enable=en)
 new_user.update_attribute("displayName",name) #same as above
 new_user.update_attribute("mail",mail)
 new_user.update_attribute("userPrincipalName",un+"@"+domain)
 new_user.update_attribute("sAMAccountName",un)
 new_user.update_attribute("description",desc)
 new_user.update_attribute("profilePath","\\\\"+domain+"\\dfs\\Profiles\\"+un)
 return

def update( user1,desc1,mail1):
 "update attributes (if user in AD and in LA)"
 try: 
  user = pyad.aduser.ADUser.from_cn(user1)
  #user.clear_attribute("mail")
  user.update_attribute("description",desc1)
  user.append_to_attribute("mail", mail1)
 except Exception as e:
  "1";  
  
 return 
   
def move( user1,enabled):
 "Move (if user in AD and not in LA) to reunisten if user has been enabled before else to lid af"
 
 # Acc enabled or not (userAccountControl)
 # 512=Enabled
 # 514= Disabled
 # 66048 = Enabled, password never expires
 # 66050 = Disabled, password never expires
 
 user = pyad.aduser.ADUser.from_cn(user1,ou_base)
 try:
  if enabled == 514 or 66050:
   ou = pyad.adcontainer.ADContainer.from_dn(ou_disabled)
   user.move(ou)
  
  if enabled == 512 or 66048:
   ou = pyad.adcontainer.ADContainer.from_dn(ou_reunisten)
   user.move(ou)   
     
 except Exception as e:
  print e
 
 return


 
# 
# Set functions  
#

 
#create set
s=set(la_users_name)
t=set(ad_users_name)

#do set
c = s - t #create based on LA - new member
m = t - s #move based on AD - old member
e = s & t #edit based on LA - current member

#create list from set
list_create=list(c)
list_move=list(m)
list_edit=list(e)

#display stats
print "Created: ",len(c)
print "Moved: ",len(m)
print "Edited: ",len(e)


#
# Execute functions after question (ctrl c to quit)
#


raw_input("Continue? ") 
for a in la_users:
 for b in list_create:
  if b==a[0]: 
   try:
	create(a[0],a[1],a[2],a[3])
   except Exception as e:
	print e
	
 for c in list_edit:
  if c==a[0]: 
   try:
	update(a[0],a[2],a[3])
   except Exception as e:
	print e	

for a in ad_users:
 for b in list_move:
  if b==a[0]: 
   try:
	move(a[0],a[4])
   except Exception as e:
	print e
	
