Introduction
============
Using a CSV-file and pyad (https://github.com/zakird/pyad) to check AD and either create, move, edit or enable objects.

This script will create four sets: 
csv not in AD = create,
ad not in csv = move,
in AD and in csv = edit,
disabled users in AD = enable

Create a user
Move the user to an other OU (for example, Disabled User)
Update values for some attributes (like mail and phone)
Enable users when disabled.

Requirements
============

Windows Server 2008r2 or better

Python 2.7.x

    python -m pip install {modules}
    ConfigParser
    csv
    smtplib
    email.MIMEMultipart
    email.MIMEText
    email.mime.application 
    pyad
    
pyad requires pywin32, available at http://sourceforge.net/projects/pywin32.

Create folder with Python_script_AD.py, config.cfg and two folders called 'logs' and 'reports'

Fill in config.cfg with desired values

Usage
============
In CMD

    python Python_script_AD.py
    
Parameters:

    Python_script_AD.py -m [cema] -d [0,1] -r [0,1,2] -f
    
(m)ode: (c)reate,(e)dit,(m)ove,(a)ctivate

(d)ebug: 0=console+log 1=log

(r)eport: 0=none, 1=local, 2=local+mail 

(f)orce ignore warning

Default recommended usage:

    Python_script_AD.py -m cem -d 0 -r 2 
    
