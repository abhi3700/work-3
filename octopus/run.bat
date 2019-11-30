@echo off
del /f octopus.dot octopus.png
python octopus.py
rem dot octopus.dot -T png -o octopus.png
rem pause