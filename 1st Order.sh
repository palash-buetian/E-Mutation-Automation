#!/bin/bash

clear 

source ~/Documents/www/E-Mutation\ Automation/venv/bin/activate

# venv is now active
cd ~/Documents/www/E-Mutation\ Automation/

echo '1st orders processing....'

python 1st_order.py

now=$(date +"%d-%m-%Y_%I.%M-%p");

FILE='1st_orders/'${now}'_1st_Order.xlsx'

#echo $FILE

if [ -f "$FILE" ]; then
	echo 'Printing Summary...'
	# add a printing  que
	lpr $FILE
else 
	echo "Nothing to print."
fi

echo -n Press enter to end...
read x
echo Ending Now\!