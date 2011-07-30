#!/bin/bash
export OO=libreoffice
export LD_LIBRARY_PATH=$LD_LIBRARY_PATH:/usr/lib/$OO/program/
export PYTHONPATH=$PYTHONPATH:/usr/lib/$OO/program
#export DISPLAY=localhost:1
export DISPLAY=:0
echo "Logging en pyUnoServer.log"
python pyUnoServerV2.py
echo "Ciao!"
