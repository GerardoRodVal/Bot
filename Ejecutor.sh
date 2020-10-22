#!/bin/bash

#git lfs checkout
#git lfs fetch 

echo Ejecutando requirements.txt
pip3 install -r requirements.txt
echo Requeriments instalados

echo Ejecutando Bot
python3 Bot_UNESCO.py
