# AD-Pyad
Using CSV and pyad (https://github.com/zakird/pyad) check AD and either create, move or edit objects.

This script will first loop over the values from AD and either move the user to "disabled OU" or update values for some attributes
Then it will loop over the values from the CSV and if necessairy create a user

The best option will be to use sets however it is not working now:


ad not in csv -> move

Intersection (csv,ad) -> Update

csv not in ad -> create
