#!/bin/bash

clear 

source ~/Documents/www/E-Mutation\ Automation/venv/bin/activate

# venv is now active
cd ~/Documents/www/E-Mutation\ Automation/

echo '2nd orders processing....'

python 2nd_order.py

echo -n Press enter to end...
read x
echo Ending Now\!