#!/bin/sh

mkdir -p masher
cp jdcal.py masher/
cp -r openpyxl masher/
cp sheetmasher.py masher/
find masher/ -name "*.pyc" -type f -delete
find masher/ -name "__pycache__" -type d -delete
zip -r masher masher
scp masher.zip tubbs:/var/www/default/

