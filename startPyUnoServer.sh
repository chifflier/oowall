#!/bin/bash
export LD_LIBRARY_PATH=$LD_LIBRARY_PATH:/usr/lib/openoffice/program/
export PYTHONPATH=$PYTHONPATH:/usr/lib/openoffice/program
#export DISPLAY=localhost:1
export DISPLAY=:0
echo "Logging en pyUnoServer.log"
python pyUnoServerV2.py
echo "Ciao!"
