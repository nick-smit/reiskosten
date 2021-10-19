# Reiskosten script
This script will create a 'reiskosten declaratie' file for my current employer. It works by detecting the SSIDs of WiFi networks to determine if the user is at the office. If one of the SSIDs is detected it will write the date to a file. The file with the dates will be read by another script to send an email containing the 'reiskosten declaratie' file.

## How to use
1. Install python3 and pip
2. Run `pip3 install -r requirements.txt` in the repository directory
3. Copy the configexample.py file to config.py and fill in the blanks
4. Run `ln -s /System/Library/PrivateFrameworks/Apple80211.framework/Versions/Current/Resources/airport /usr/local/bin/airport` to make airport available to the script
5. Add the following commands to your crontab (don't forget to change the path in the cronjobs)
6. Test if the commands work.

```
# Reiskosten declaratie
0 * * * * /usr/local/bin/python3 /path/to/reiskosten/reiskosten-detect.py
5 * * * * /usr/local/bin/python3 /path/to/reiskosten/send-reiskosten.py
```
