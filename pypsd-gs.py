from sys import exit
import time
import datetime
import os
import pytz
try:
    import requests
    import gspread
    import pandas as pd
except:
    print("Python Request module not available on this machine")
    print("Fatal Error: The program will now quit!")
    print("Run <pip install requests> to install this module ")
    exit()
import re
from platform import platform

class bcolors:
    HEADER = '\033[95m'
    OKGREEN = '\033[92m'
    OKBLUE = '\033[94m'
    OKPURPLE = '\033[95m'
    INFOYELLOW = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'

try:
    psdemail=os.environ["psdemail"]
    psdpass=os.environ["psdpass"]
    
except:
    exit(f"{bcolors.FAIL}Input Email and Password as Environment Variables{bcolors.ENDC}")

try:
    credentials={
  "type": "service_account",
  "project_id": os.environ["project_id"],
  "private_key_id": os.environ["private_key_id"],
  "private_key": os.environ["private_key"].replace('\\n','\n'),
  "client_email": os.environ["client_email"],
  "client_id": os.environ["client_id"],
  "auth_uri": "https://accounts.google.com/o/oauth2/auth",
  "token_uri": "https://oauth2.googleapis.com/token",
  "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
  "client_x509_cert_url": os.environ["client_x509_cert_url"],
  "universe_domain": "googleapis.com"
}
except:
    exit(f"{bcolors.FAIL}Error loading credentials from Environment Variables{bcolors.ENDC}")

try:
    sheetlink=os.environ["sheetlink"]
except:
    exit(f"{bcolors.FAIL}Error sheetlink from Environment Variables{bcolors.ENDC}")

scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

client = gspread.service_account_from_dict(credentials,scopes=scopes)

wb = client.open_by_url(sheetlink)

studentid=0

projectlist='2023-2024 / SEM-I' #project list for which data will be fetched. Entire history is sent by the server thus has to be filtered

if (platform()[0:7]=="Windows"):
    import ctypes
    kernel32 = ctypes.windll.kernel32
    kernel32.SetConsoleMode(kernel32.GetStdHandle(-11), 7)

def token_gen():
    curr_millis=int(time.time()*1000)
    epochTicks = 621355968000000000
    ticksPerMillisecond = 10000
    token= epochTicks + (curr_millis * ticksPerMillisecond)
    return int(token)


f=open('debug','w')

url='http://psd.bits-pilani.ac.in/Login.aspx'

resp_url=f'http://psd.bits-pilani.ac.in/Student/NEWStudentDashboard.aspx?StudentId={studentid}'

resp_url2='http://psd.bits-pilani.ac.in/Student/StudentStationPreference.aspx'

station_fetch='http://psd.bits-pilani.ac.in/Student/ViewActiveStationProblemBankData.aspx'

pb_details='http://psd.bits-pilani.ac.in/Student/StationproblemBankDetails.aspx'

ps=requests.Session()
resp_get=ps.get(url)
outhtml = resp_get.text.splitlines()
test=0
print(f"{bcolors.OKBLUE}>{bcolors.ENDC}Getting Validation token...")
formelement=[]
for line in outhtml:
    if (line.find('<form')>0):
        test=test+1
    if test>0:
        formelement.append(line)
    if (line.find('</form')>0):
        test=test-1

inputelement=[]
for line in formelement:
    if (line.find('VIEWSTATE')>0):
        inputelement.append(line)
    if (line.find('EVENTVALIDATION')>0):
        inputelement.append(line)
inval=[]
for i in range(len(inputelement)):
    inputelement[i]=inputelement[i].split(' ')
    for entry in inputelement[i]:
        if (len(entry)>5 and entry[0:5]=='value'):
            inputelement[i]=entry[7:-1]

payload = {
    '__EVENTTARGET':"",
    '__EVENTARGUMENT':"",
    '__VIEWSTATE':inputelement[0],
    '__VIEWSTATEGENERATOR':inputelement[1],
    '__EVENTVALIDATION':inputelement[2],
    'TxtEmail': psdemail,
    'txtPass': psdpass,
    'Button1':"Login",
    'txtEmailid':""
}

payload2 ={'CompanyId':"0"}

headers={
    'Host': 'psd.bits-pilani.ac.in',
    'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64; rv:109.0) Gecko/20100101 Firefox/113.0',
    'Accept': 'application/json, text/javascript, */*; q=0.01',
    'Accept-Language': 'en-US,en;q=0.5',
    'Accept-Encoding': 'gzip, deflate',
    'Content-Type': 'application/json; charset=utf-8',
    'X-Requested-With': 'XMLHttpRequest',
    'Origin': 'http://psd.bits-pilani.ac.in',
    'Connection': 'keep-alive',
    'Referer': 'http://psd.bits-pilani.ac.in/Student/StudentStationPreference.aspx'
}
print(f"{bcolors.OKBLUE}>{bcolors.ENDC}Logging in...")
wb.sheet1.update([['Logging in Please Wait...']])
post_req=ps.post(url,data=payload)
outhtml=post_req.text.splitlines()

for line in outhtml:
    if (line.find('StudentId')>0):
        studentid=line.split(' ')

try:
    for entry in studentid:
        if (entry[-16:-7]=='StudentId'):
            studentid=entry[-6:-1]
    print(f"{bcolors.OKGREEN}SUCCESS{bcolors.ENDC}\n")
except:
    wb.sheet1.merge_cells("A1:G2",merge_type='MERGE_ROWS')
    curr_time=datetime.datetime.now(pytz.timezone('Asia/Kolkata'))
    wb.sheet1.update([['Writing Sheet Please Wait...']])
    wb.sheet1.update([['Sheet automatically updated using Github Actions + pypsd_bot(github.com/byethon/psms-to-gsheet)']]+[[f'LOGIN DISABLED OR EXECUTION ERROR : Next update at {(curr_time+datetime.timedelta(hours=1)).strftime("%b %d %Y %H:%M%p")}']],value_input_option="USER_ENTERED")
    wb.sheet1.format(f"A1:G1", {"textFormat": {{'bold': True},{"foregroundColor": {"red": 0.4,"green": 0.4,"blue": 0.4}}}})
    wb.sheet1.format(f"A2:G2", {"textFormat": {{'bold': True},{"foregroundColor": {"red": 0.92,"green": 0.26,"blue": 0.21}}}})
    exit(f"{bcolors.FAIL}Check Email and Password{bcolors.ENDC}")

get_req=ps.get(resp_url)
print(f"{bcolors.OKBLUE}>{bcolors.ENDC}Getting Dashboard...")
get_req=ps.get(resp_url2)
print(f"{bcolors.OKBLUE}>{bcolors.ENDC}Getting Station List....")
wb.sheet1.update([['Getting Station List...']])
post_req=ps.post(resp_url2+'/getinfoStation',headers=headers,json=payload2)

if(post_req.status_code==404):
    print(f"{bcolors.FAIL}ERROR accessing Station Preference list{bcolors.ENDC}")
    print(f"{bcolors.OKBLUE}>{bcolors.ENDC}Fallback to problem bank scraping....")
    print(f"{bcolors.OKBLUE}>{bcolors.ENDC}Getting Station List....")
    payload_bak={'batchid': "undefined",
                'token': str(token_gen())}

    headers.update({'Referer': station_fetch})
    post_req=ps.post(station_fetch+'/getPBdetail',headers=headers,json=payload_bak)

    jsonout=post_req.text[8:-4]
    jsonout=jsonout.split('},{')
    for i in range(len(jsonout)):
        jsonout[i]=re.sub('\\\\"','',jsonout[i])
        jsonout[i]=re.sub('\\\\\\\\u0026','&',jsonout[i])
        jsonout[i]=jsonout[i].split(',')
        temp=[]
        for subentry in jsonout[i]:
            temprep=subentry.split(':',1)
            if len(temprep)>1:
                temp.append(temprep[1].rstrip())
            else:
                temp[-1]=temp[-1]+','+subentry.rstrip()
        jsonout[i]=temp
    for i in range(len(jsonout)):
        temp=jsonout[i][0]
        jsonout[i][0]=jsonout[i][-2]
        jsonout[i][-2]=temp
        temp=jsonout[i][6]
        jsonout[i][6]=jsonout[i][-1]
        jsonout[i][-1]=temp
        temp=jsonout[i][2]
        jsonout[i][2]=jsonout[i][-6]+'-'+jsonout[i][-3]+', '+jsonout[i][-4]
        jsonout[i][-3]=temp
else:
    jsonout=post_req.text[8:-4]
    jsonout=jsonout.split('},{')
    for i in range(len(jsonout)):
        jsonout[i]=re.sub('\\\\"','',jsonout[i])
        jsonout[i]=re.sub('\\\\\\\\u0026','&',jsonout[i])
        jsonout[i]=jsonout[i].split(',')
        temp=[]
        for subentry in jsonout[i]:
            temprep=subentry.split(':',1)
            if len(temprep)>1:
                temp.append(temprep[1].rstrip())
            else:
                temp[-1]=temp[-1]+','+subentry.rstrip()
        jsonout[i]=temp

print(f"{bcolors.OKGREEN}RECIEVED{bcolors.ENDC}\n")

for j in range(len(jsonout)):
    try:
        [Sdomain,StationName]=jsonout[j][2].split('-',1)
    except:
        [Sdomain,StationName]=['-',jsonout[j][2]]
    print(f'{bcolors.INFOYELLOW}>{bcolors.ENDC}{Sdomain.rstrip()}-{StationName}')

print(f"\n{bcolors.OKBLUE}>{bcolors.ENDC}Fetching Project list...\n")
headers.update({'Referer': station_fetch})
fetchlist=[]
print(f"{bcolors.OKBLUE}>{bcolors.ENDC}Fetching data....")
wb.sheet1.update([['Fetching Project list...']])
for entry in jsonout:
    payload3={
        'StationId':f'{entry[-2]}'
    }
    post_req=ps.post(station_fetch+'/getPBPOPUP',headers=headers,json=payload3)
    fetchlist.append(re.sub('\\\\"','',post_req.text[8:-4]))
    fetchlist[-1]=fetchlist[-1].split('},{')
    for i in range(len(fetchlist[-1])):
        fetchlist[-1][i]=fetchlist[-1][i].split(',')
        temp=[]
        for subentry in fetchlist[-1][i][2:]:
            temprep=subentry.split(':',1)
            if len(temprep)>1:
                temp.append(temprep[1].rstrip())
            else:
                temp[-1]=temp[-1]+','+subentry.rstrip()
        fetchlist[-1][i]=temp

print(f"{bcolors.OKGREEN}RECIEVED{bcolors.ENDC}\n")
print(f"{bcolors.OKBLUE}>{bcolors.ENDC}Filtering for incomplete data and Stripend Constraints")
wb.sheet1.update([['Filtering for incomplete/broken data and Past Semester Data...']])
print(f"{bcolors.FAIL}ALSO REMOVING PAST SEMESTER DATA{bcolors.ENDC}")
pop_arroext=[]
for i in range(len(fetchlist)):
    pop_arrint=[]
    for j in range(len(fetchlist[i])):
        fetchlist[i][j].append(j)
        if(len(fetchlist[i][j])<2 or (fetchlist[i][j][1]!=projectlist and j)):
            if(len(fetchlist[i][j])<2):
                print(f"{bcolors.INFOYELLOW}>{bcolors.ENDC}No Fetch File? Check Debug Output")
                f.write(str(jsonout[i]))
                f.write('-------------')
                f.write(str(fetchlist[i]))
                f.write('\n')
            pop_arrint.append(j)
            try:
                print(f"{bcolors.INFOYELLOW}>{bcolors.ENDC}Project List {bcolors.OKBLUE}{fetchlist[i][j][1]}{bcolors.ENDC} removed-{jsonout[i][2]}")
            except:
                print(f"{bcolors.INFOYELLOW}>{bcolors.ENDC}Removed Broken list data")
    pop_arrint.reverse()
    for entry in pop_arrint:
        fetchlist[i].pop(entry)
    if(len(fetchlist[i])==0):
        pop_arroext.append(i)
        print(f"{bcolors.INFOYELLOW}>{bcolors.ENDC}Station removed-{bcolors.FAIL}{jsonout[i][2]}{bcolors.ENDC}")

pop_arroext.reverse()
for entry in pop_arroext:
    fetchlist.pop(entry)
    jsonout.pop(entry)
print(f"{bcolors.OKGREEN}DONE{bcolors.ENDC}\n")

print(f"{bcolors.OKBLUE}>{bcolors.ENDC}Generating Project Sublist URLS")
for i in range(len(jsonout)):
    try:
        [Sdomain,StationName]=jsonout[i][2].split('-',1)
    except:
        [Sdomain,StationName]=['-',jsonout[i][2]]
    for j in range(len(fetchlist[i])):
        urlmask=f'http://psd.bits-pilani.ac.in/Student/StationproblemBankDetails.aspx?CompanyId={jsonout[i][-1]}&StationId={jsonout[i][-2]}&BatchIdFor={fetchlist[i][j][2]}&PSTypeFor={fetchlist[i][j][3]}'
        fetchlist[i][j].append(urlmask)

payload4={"batchid": "undefined"}

Stationcol=[]
Domcol=[]
Lupdcol=[]
Eligcol=[]
TotalReqcol=[]
Stripcol=[]
Linkcol=[]
print(f"{bcolors.OKBLUE}>{bcolors.ENDC}Fetching Project Sublists and Generating Output")
print(f"{bcolors.OKBLUE}>{bcolors.ENDC}Fetching data....")
wb.sheet1.update([['Fetching Project details and Generating Output...']])
for i in range(len(jsonout)):
    try:
        [Sdomain,StationName]=jsonout[i][2].split('-',1)
    except:
        [Sdomain,StationName]=['-',jsonout[j][2]]
    Stationcol.append(StationName)
    Domcol.append(Sdomain)
    for j in range(len(fetchlist[i])):
        uri=fetchlist[i][j][-1]
        headers.update({'Referer': uri})
        post_req=ps.get(uri)
        post_req=ps.post(pb_details+'/ViewPB',headers=headers,json=payload4)
        pbout=post_req.text[8:-3]
        pbout=pbout.split('},{')
        for k in range(len(pbout)):
            pbout[k]=re.sub('\\\\"','',pbout[k])
            pbout[k]=re.sub('\\\\\\\\u0026','&',pbout[k])
            pbout[k]=pbout[k].split(',')
            temp=[]
            for subentry in pbout[k]:
                temprep=subentry.split(':',1)
                if len(temprep)>1:
                    temp.append(temprep[1].rstrip())
                else:
                    temp[-1]=temp[-1]+','+subentry.rstrip()
            pbout[k]=temp
        totalinterns=0
        last_updated=''
        for k in range(len(pbout)):
            valid=False
            totalinterns=totalinterns+int(pbout[k][1])
            last_updated=pbout[k][-4]
        Lupdcol.append(last_updated)
        Eligcol.append(fetchlist[i][j][-6])
        TotalReqcol.append(totalinterns)
        Stripcol.append(fetchlist[i][j][-5])
        Linkcol.append(f'=HYPERLINK("{uri}","View Details")')

dataset = {
  'Station': Stationcol,
  'Domain': Domcol,
  'Last updated on':Lupdcol,
  'Eligibility':Eligcol,
  'Total Req. Interns':TotalReqcol,
  'Stripend':Stripcol,
  'Link':Linkcol
}
dataframe=pd.DataFrame(dataset)
dataframe['Last updated on']=pd.to_datetime(dataframe['Last updated on'], format='%b  %d %Y  %H:%M%p')
dataframe.sort_values(by='Last updated on',ascending=False, inplace=True)
dataframe.style.format({"Last updated on": lambda t: t.strftime("%b  %d %Y  %H:%M%p")})
dataframe['Last updated on']=dataframe['Last updated on'].dt.strftime('%b %d %Y %H:%M%p')
wb.sheet1.clear()
wb.sheet1.update([['Writing Sheet Please Wait...']])
wb.sheet1.merge_cells("A1:G2",merge_type='MERGE_ROWS')
curr_time=datetime.datetime.now(pytz.timezone('Asia/Kolkata'))
wb.sheet1.update([['Sheet automatically updated using Github Actions + pypsd_bot(github.com/byethon/psms-to-gsheet)']]+[[f'Sheet Last updated at {curr_time.strftime("%b %d %Y %H:%M%p")} next update at {(curr_time+datetime.timedelta(hours=1)).strftime("%b %d %Y %H:%M%p")}']]+[dataframe.columns.values.tolist()] + dataframe.values.tolist(),value_input_option="USER_ENTERED")
wb.sheet1.format("A3:G3",{'textFormat': {'bold': True}})
wb.sheet1.freeze(rows=3)
row_count=wb.sheet1.row_count
wb.sheet1.format(f"A1:G2", {"textFormat": {{'bold': True},{"foregroundColor": {"red": 0.4,"green": 0.4,"blue": 0.4}}}})
wb.sheet1.format(f"A4:A{row_count}", {"textFormat": {"foregroundColor": {"red": 0.6,"green": 0.0,"blue": 1.0}}})
wb.sheet1.format(f"F4:F{row_count}", {"textFormat": {"foregroundColor": {"red": 0.2,"green": 0.66,"blue": 0.33}}})
#wb.sheet1.format(f"F4:F{row_count}", {"textFormat": {"foregroundColor": {"red": 0.07,"green": 0.34,"blue": 0.8}}})
wb.sheet1.format(f"E4:E{row_count}", {"textFormat": {"foregroundColor": {"red": 1.0,"green": 0.43,"blue": 0.1}}})
wb.sheet1.format(f"C4:C{row_count}", {"textFormat": {"foregroundColor": {"red": 0.75,"green": 0.56,"blue": 0.0}}})
print("Program executed Successfuly")

