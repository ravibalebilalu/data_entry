#!/bin/bash
cd ..
target_directory="/data_entry"
source_directory="data_entry"
cp -r $source_directory /
cd /$target_directory
echo "diectory created"

echo "Things are getting ready, lean back"
virtualenv -p python3 de
echo "venv created"

source de/bin/activate
echo "venv activated"
echo "Installing required packages...."
pip install --upgrade pip
pip install -r requirements.txt
deactivate
cp excel.sh /usr/local/bin/excel


result=$?
if [ $result -eq 0 ];then
    echo "Installation compleated succusfully!"
    echo "Type 'excel' and hit enter to run the program"
else
    echo "Something not working properly"
fi


 
