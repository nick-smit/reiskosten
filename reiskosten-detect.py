from os import path
import subprocess
import re
import json
import datetime
import config

def getAvailableSSIDS() -> list:
    print("Fetching available ssids")
    # using the check_output() for having the network term retrival
    result = subprocess.run(['/usr/local/bin/airport', '-s'], stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
    scan_err = result.stderr
    scan_out = result.stdout

    if scan_err != None:
        print(f"Unable to scan for wifi networks: {scan_err}")
        return []

    scan_out_lines = str(scan_out).split("\\n")[1:-1]

    ssids = []
    for each_line in scan_out_lines:
        ssids.append(re.split('\s([0-9a-f]{2}:?){6}\s', each_line)[0].strip())

    print(f"Found ssids: {', '.join(ssids)}")

    return ssids

def isAtOffice():
    available_ssids = getAvailableSSIDS()

    for ssid in config.detect_ssids:
        if ssid in available_ssids:
            return True

    return False

if __name__ == '__main__':
    try:
        print("Executing reiskosten-detect.py")
        today = datetime.date.today()

        # after 20th should be next month
        month = today.month
        year = today.year
        if today.day > config.send_on_day_of_month:
            if month == 12:
                # add 1 for december case.
                year += 1
            month = (month + 1) % 12

        
        curFile = f"{config.base_dir}data/{today.year}-{month}.json"
        data = []
        if not path.isfile(curFile):
            with open(curFile, 'w') as file:
                file.write('[]\n')
        else:
            with open(curFile, 'r') as file:
                data = json.loads(file.read())

        if today.isoformat() in data:
            print("Already written to file")
            exit(0)
        
        if isAtOffice():
            print("Welcome to the office!")
            data.append(today.isoformat())

            with open(curFile, 'w') as file:
                file.write(json.dumps(data) + "\n")
        else:
            print("Didn't detect any office WiFi signals")
    except Exception as e:
        print(f"Caught exception: {e}")
