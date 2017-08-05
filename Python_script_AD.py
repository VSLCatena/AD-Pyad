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

ou_base = config.get('ou', 'base')
ou_disabled = config.get('ou','disabled')

user_pw=config.get('user','pw')
user_domain=config.get('user','domain')


#connect to AD
#
#read CSV
#
# ['name','username','descr',email']
#
import csv
c=0
la_users=[]
with open('FILENAME.CSV', 'rb') as f:
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
    attributes = [ "description", "displayName", "mail", "sAMAccountName"],
    where_clause = "objectClass = 'user'",
    base_dn = ou_base
)
ad_users = []
for x in q.get_results():
 ad_user_data= []
 ad_user_data.append(x["displayName"])
 ad_user_data.append(x["sAMAccountName"])
 ad_user_data.append(x["mail"])
 ad_users.append(ad_user_data)
ad_users_name=[xx[0] for xx in ad_users]
#print ad_users_name

 
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
   
def move( user1):
 "Move (if user in AD and not in LA)"
 user = pyad.aduser.ADUser.from_cn(user1,ou_base)
 try:
  ou_disabled = pyad.adcontainer.ADContainer.from_dn(ou_disabled)
  user.move(ou_disabled)
 except Exception as e:
  print e
 #ERROR: (-2147352567, 'Exception occurred.', (0, u'Active Directory', u'There is no such object on the server.\r\n', None, 0, -2147016656), None)
 return


 
# 
# Set functions  
#
 
#create set
s=set(la_users_name)
t=set(ad_users_name)

#do set
c = s - t #create based on LA
m = t - s #move based on AD
e = s & t #edit based on LA

#create list from set
list_create=list(c)
list_move=list(m)
list_edit=list(e)

#display stats
print "Created: ",len(c)
print "Moved: ",len(m)
print "Edited: ",len(e)

# for y in list_move:
# move_disabled(y)



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
    edit(a[0],a[2],a[3])
   except Exception as e:
    print e	

for a in ad_users:
 for b in list_move:
  if b==a[0]: 
   try:
    move(a[0])
   except Exception as e:
    print e
	
