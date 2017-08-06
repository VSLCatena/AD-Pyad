# AD-Pyad
Using a CSV and pyad (https://github.com/zakird/pyad) to check AD and either create, move or edit objects.

This script will create three sets: 
csv not in ad = create,
ad not in csv = move,
in ad and in csv = edit,

and either move the user to "disabled OU" or update values for some attributes and if necessairy create a user
