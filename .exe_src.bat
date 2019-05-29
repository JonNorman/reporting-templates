echo OFF
mode con: cols=280 lines=80
echo Checking module dependencies...
echo ON
pip install -r requirements.txt
echo OFF
echo Dependencies successfully loaded; running Report Formatter
python .\src\reporter.py
pause
