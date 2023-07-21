from sys import exit
import time
import datetime
import os
import re
from multiprocessing import Process,Queue,set_start_method
from platform import platform
env_file = os.getenv('GITHUB_ENV')
try:
    import pytz
    import requests
    import gspread
    import pandas as pd
    import aiohttp
    import asyncio
except:
    with open(env_file, "a") as myfile:
        myfile.write("RETRY_PY=1")
    print("All required modules not available on this machine")
    print("Fatal Error: The program will now quit!")
    print("Retrying...")
    time.sleep(2)
    exit()

REQUEST_THREADS=36 #No. of threads from which to send server requests (More Threads are faster but performance saturates at some point and drops beyond it)
RETRY_COUNT=5
studentid=0

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

def put_login_queue(queue):
    queue.put(gen_login_session())

def gen_login_multi(Login_threads):
    session_list=[]
    q1=Queue()
    p=[]
    for k in range(Login_threads):
            p.append(Process(target = put_login_queue,args=(q1, )))
            p[k].start()
    for k in range(Login_threads):
        Req_out=q1.get()
        if(Req_out[0] != None):
            session_list.append(Req_out[1])
    for k in range(Login_threads):
        p[k]=Process(target=keep_valid_logins,args=(k,session_list[k],q1))
        p[k].start()
    for k in range(Login_threads):
        p[k].join()
    rem_session=[]
    while not q1.empty():
        rem_session.append(q1.get())
    rem_session=rem_session.sort(reverse=True)
    if rem_session:
        for entry in rem_session:
            session_list.pop(entry)
    q1.close()
    return session_list

def keep_valid_logins(k,requests_session,q1):
    if(not login_test(requests_session)):
        q1.put(k)


def gen_login_session():
    url='http://psd.bits-pilani.ac.in/Login.aspx'
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
    post_req=ps.post(url,data=payload)
    outhtml=post_req.text.splitlines()

    for line in outhtml:
        if (line.find('StudentId')>0):
            studentid=line.split(' ')

    try:
        for entry in studentid:
            if (entry[-16:-7]=='StudentId'):
                studentid=entry[-6:-1]
        print(f"{bcolors.OKGREEN}Getting Student ID and Logging in...{bcolors.ENDC}\n")
    except:
        print("Login Error")
        return [0,None]
    return [studentid,ps]

def login_test(requests_session):
    resp_url=f'http://psd.bits-pilani.ac.in/Student/NEWStudentDashboard.aspx?StudentId={studentid}'
    if(requests_session.head(resp_url).status_code==200):
        print(f"{bcolors.OKGREEN}Login Valid{bcolors.ENDC}")
        return 1
    else:
        print(f"{bcolors.FAIL}Login Invalid{bcolors.ENDC}")
        return 0

def token_gen():
    curr_millis=int(time.time()*1000)
    epochTicks = 621355968000000000
    ticksPerMillisecond = 10000
    token= epochTicks + (curr_millis * ticksPerMillisecond)
    return int(token)

async def collector(session,url,payload,headers=''):
    async with session.post(url,json=payload,headers=headers) as resp:
        respout=await resp.text()
        return respout

async def detail_collector(cookies,url,headers,Stationlist,resp_list):
    async with aiohttp.ClientSession(cookies=cookies,headers=headers) as psasync:
        tasks=[]
        for i in range(len(Stationlist)):
            payload3={
            'StationId':f'{Stationlist[i][-2]}'
        }
            tasks.append(asyncio.ensure_future(collector(psasync,url,payload3,headers)))
        
        orig_resp= await asyncio.gather(*tasks)
        resp_list.extend(orig_resp)
        
        
def detail_fetchv2(requests_session,Stationlist,url,headers,queue):
    cookies=requests_session.cookies.get_dict()
    Proj_list=[]
    resp_list=[]
    Errorcount=0
    while(Errorcount>=0 and Errorcount<RETRY_COUNT):
        try:
            asyncio.run(detail_collector(cookies,url+'/getPBPOPUP',headers,Stationlist,resp_list))
            Errorcount=-1
        except:
            print(f"{bcolors.FAIL}>{bcolors.ENDC}Error Encountered Retrieving data")
            Errorcount+=1
            if(Errorcount<RETRY_COUNT):
                print("Retrying...")
    if(Errorcount>=RETRY_COUNT):
        queue.put([[],[]])
        exit("Too many errors recieving data")
    for j in range(len(resp_list)):
        Proj_list.append(re.sub('\\\\"','',resp_list[j][8:-4]))
        Proj_list[-1]=Proj_list[-1].split('},{')
        for i in range(len(Proj_list[-1])):
            Proj_list[-1][i]=Proj_list[-1][i].split(',')
            temp=[]
            for subentry in Proj_list[-1][i][2:]:
                temprep=subentry.split(':',1)
                if len(temprep)>1:
                    temp.append(temprep[1].rstrip())
                else:
                    temp[-1]=temp[-1]+','+subentry.rstrip()
            Proj_list[-1][i]=temp
    
    queue.put([Stationlist,Proj_list])

def proj_fetch(request_session,Stationlist,Proj_list,fetch_url,headers,payload,queue):
    Stationcol=[]
    Domcol=[]
    Loccol=[]
    Lupdcol=[]
    Addcol=[]
    Eligcol=[]
    TotalReqcol=[]
    Stripcol=[]
    Linkcol=[]
    for i in range(len(Stationlist)):
        try:
            [Sdomain,StationName]=Stationlist[i][2].split('-',1)
        except:
            [Sdomain,StationName]=['-',Stationlist[i][2]]
        [StationName,location]=StationName.rsplit(',',1)
        StationName=StationName.strip()
        Sdomain=Sdomain.strip()
        location=location.strip()
        Loccol.append(location)
        Stationcol.append(StationName)
        Domcol.append(Sdomain)
        for j in range(len(Proj_list[i])):
            uri=Proj_list[i][j][-1]
            headers.update({'Referer': uri})
            Errorcount=0
            while(Errorcount>=0 and Errorcount<RETRY_COUNT):
                try:
                    post_req=request_session.head(uri)
                    post_req=request_session.post(fetch_url+'/ViewPB',headers=headers,json=payload)
                    print(f'{bcolors.INFOYELLOW}>{bcolors.ENDC}{Sdomain.rstrip()}-{StationName}')
                    pbout=post_req.text[8:-3]
                    pbout=pbout.split('},{')
                    totalinterns=0
                    last_updated=''
                    added_on=''
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
                        totalinterns=totalinterns+int(pbout[k][1])
                        pbout[k][-4]=datetime.datetime.strptime(pbout[k][-4], '%b  %d %Y  %H:%M%p')
                        if(k==0 or pbout[k][-4]>last_updated):
                            last_updated=pbout[k][-4]
                        if(k==0 or pbout[k][-4]<added_on):
                            added_on=pbout[k][-4]
                    Lupdcol.append(last_updated)
                    Addcol.append(added_on)
                    Eligcol.append(Proj_list[i][j][-6])
                    TotalReqcol.append(totalinterns)
                    Stripcol.append(Proj_list[i][j][-5])
                    Linkcol.append(f'=HYPERLINK("{uri}","View Details")')
                    Errorcount=-1
                except:
                    print(f"{bcolors.FAIL}>{bcolors.ENDC}Error Encountered Retrieving data for {Sdomain.rstrip()}-{StationName}")
                    Errorcount+=1
                    if(Errorcount<RETRY_COUNT):
                        print("Retrying...")
            if(Errorcount>=RETRY_COUNT):
                print(f"{bcolors.FAIL}>{bcolors.ENDC}Error Encountered Retrieving data for Station {Sdomain.rstrip()}-{StationName}, Skipping")
                break
            
    queue.put([Stationcol,Domcol,Loccol,Lupdcol,Addcol,Eligcol,TotalReqcol,Stripcol,Linkcol])


if __name__=='__main__':
    set_start_method("spawn")

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

    projectlist=''           #'2023-2024 / SEM-I' #project list for which data will be fetched. Entire history is sent by the server thus has to be filtered

    if (platform()[0:7]=="Windows"):
        import ctypes
        kernel32 = ctypes.windll.kernel32
        kernel32.SetConsoleMode(kernel32.GetStdHandle(-11), 7)

    f=open('debug','w')

    resp_url=f'http://psd.bits-pilani.ac.in/Student/NEWStudentDashboard.aspx?StudentId={studentid}'


    resp_url2='http://psd.bits-pilani.ac.in/Student/StudentStationPreference.aspx'

    station_fetch='http://psd.bits-pilani.ac.in/Student/ViewActiveStationProblemBankData.aspx'

    pb_details='http://psd.bits-pilani.ac.in/Student/StationproblemBankDetails.aspx'

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
    
    [studentid,ps]=gen_login_session()
    
    if(ps == None):
        exit("Login Failed")

    get_req=ps.head(resp_url)
    print(f"{bcolors.OKBLUE}>{bcolors.ENDC}Getting Dashboard...")
    get_req=ps.head(resp_url2)
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

    print(f"\n{bcolors.OKBLUE}>{bcolors.ENDC}Fetching Project list...\n")
    headers.update({'Referer': station_fetch})

    wb.sheet1.update([['Fetching Project list...']])

    print(f"{bcolors.OKBLUE}>{bcolors.ENDC}Invoking {REQUEST_THREADS}x Parallel request threads...")
    print(f"{bcolors.FAIL}WARNING:{bcolors.ENDC}Too many threads may cause the server to disconnect!")

    login_arr=[]
    login_arr.append(ps)
    ps.close()
    login_arr.extend(gen_login_multi(REQUEST_THREADS-1))
    if(not login_test(login_arr[0])):
        login_arr.pop(0)
    REQUEST_THREADS=len(login_arr)

    if(REQUEST_THREADS == 0):
        exit("Login Unsuccessful")

    print(f"{bcolors.OKBLUE}>{bcolors.ENDC}Fetching data....")

    q1=Queue()
    Stationlist=[]
    for k in range(REQUEST_THREADS):
        Stationlist.append([])
    for z in range(len(jsonout)):
        Stationlist[z%REQUEST_THREADS].append(jsonout[z])

    fetchlist=[]
    p=[]
    jsonout=[]

    for k in range(REQUEST_THREADS):
            p.append(Process(target = detail_fetchv2,args=(login_arr[k],Stationlist[k],station_fetch,headers,q1)))
            p[k].start()
    for k in range(REQUEST_THREADS):
        Req_out=q1.get()
        fetchlist=fetchlist+Req_out[1]
        jsonout=jsonout+Req_out[0]

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
            try:
                urlmask=f'http://psd.bits-pilani.ac.in/Student/StationproblemBankDetails.aspx?CompanyId={jsonout[i][-1]}&StationId={jsonout[i][-2]}&BatchIdFor={fetchlist[i][j][2]}&PSTypeFor={fetchlist[i][j][3]}'
            except:
                print(jsonout[i])
                print(fetchlist[i])
                print(fetchlist[i][j])
                time.sleep(2)
                exit("Check recieved response")
            fetchlist[i][j].append(urlmask)

    Stationcol=[]
    Domcol=[]
    Loccol=[]
    Lupdcol=[]
    Addcol=[]
    Eligcol=[]
    TotalReqcol=[]
    Stripcol=[]
    Linkcol=[]
    print(f"{bcolors.OKBLUE}>{bcolors.ENDC}Fetching Project Sublists and Generating Output")
    print(f"{bcolors.OKBLUE}>{bcolors.ENDC}Fetching data....")
    wb.sheet1.update([['Fetching Project details and Generating Output...']])

    payload4={"batchid": "undefined"}

    Stationlist=[]
    Proj_list=[]
    for k in range(REQUEST_THREADS):
        Stationlist.append([])
        Proj_list.append([])
    for z in range(len(jsonout)):
        Stationlist[z%REQUEST_THREADS].append(jsonout[z])
        Proj_list[z%REQUEST_THREADS].append(fetchlist[z])

    for k in range(REQUEST_THREADS):
            p[k]=Process(target = proj_fetch,args=(login_arr[k],Stationlist[k],Proj_list[k],pb_details,headers,payload4,q1))
            p[k].start()
    for k in range(REQUEST_THREADS):
        Req_out=q1.get()
        Stationcol=Stationcol+Req_out[0]
        Domcol=Domcol+Req_out[1]
        Loccol=Loccol+Req_out[2]
        Lupdcol=Lupdcol+Req_out[3]
        Addcol=Addcol+Req_out[4]
        Eligcol=Eligcol+Req_out[5]
        TotalReqcol=TotalReqcol+Req_out[6]
        Stripcol=Stripcol+Req_out[7]
        Linkcol=Linkcol+Req_out[8]
    
    for k in range(REQUEST_THREADS):
        p[k].join()
        login_arr[k].close()
    q1.close()

    dataset = {
    'Station': Stationcol,
    'Location': Loccol,
    'Domain': Domcol,
    'Company Added on':Addcol,
    'Last updated on':Lupdcol,
    'Eligibility':Eligcol,
    'Total Req. Interns':TotalReqcol,
    'Stripend':Stripcol,
    'Link':Linkcol
    }
    row_count=len(Stationcol)+3
    str_list = list(filter(None, wb.sheet1.col_values(1)))
    last_row=len(str_list)
    last_col='I'
    dataframe=pd.DataFrame(dataset)
    dataframe.sort_values(by='Last updated on',ascending=False, inplace=True)
    dataframe.style.format({"Last updated on": lambda t: t.strftime("%b  %d %Y  %H:%M%p")})
    dataframe['Last updated on']=dataframe['Last updated on'].dt.strftime('%b %d %Y %H:%M%p')
    dataframe.style.format({"Company Added on": lambda t: t.strftime("%b  %d %Y  %H:%M%p")})
    dataframe['Company Added on']=dataframe['Company Added on'].dt.strftime('%b %d %Y %H:%M%p')
    if(last_row>row_count):
        wb.sheet1.batch_clear([f'A{row_count+1}:{last_col}{last_row}'])
    wb.sheet1.merge_cells("A1:I2",merge_type='MERGE_ROWS')
    wb.sheet1.format(f"A1:I2", {"textFormat": {"foregroundColor": {"red": 0.4,"green": 0.4,"blue": 0.4},'bold': True, 'underline': False}})
    wb.sheet1.format("A3:I3",{'textFormat': {'bold': True}})
    wb.sheet1.update([['Writing Sheet Please Wait...']])
    wb.sheet1.freeze(rows=3)
    wb.sheet1.format(f"A4:A{row_count}", {"textFormat": {"foregroundColor": {"red": 0.6,"green": 0.0,"blue": 1.0},'bold': True}})
    wb.sheet1.format(f"B4:B{row_count}", {"textFormat": {"foregroundColor": {"red": 0.0,"green": 0.0,"blue": 0.0}}})
    wb.sheet1.format(f"C4:C{row_count}", {"textFormat": {"foregroundColor": {"red": 0.22,"green": 0.46,"blue": 0.11},'bold': True}})
    wb.sheet1.format(f"D4:D{row_count}", {"textFormat": {"foregroundColor": {"red": 0.92,"green": 0.26,"blue": 0.21},'bold': True}})
    wb.sheet1.format(f"E4:E{row_count}", {"textFormat": {"foregroundColor": {"red": 0.75,"green": 0.56,"blue": 0.0},'bold': True}})
    wb.sheet1.format(f"F4:F{row_count}", {"textFormat": {"foregroundColor": {"red": 0.0,"green": 0.0,"blue": 0.0}}})
    wb.sheet1.format(f"G4:G{row_count}", {"textFormat": {"foregroundColor": {"red": 0.2,"green": 0.66,"blue": 0.33},'bold': True}})
    wb.sheet1.format(f"H4:H{row_count}", {"textFormat": {"foregroundColor": {"red": 0.07,"green": 0.34,"blue": 0.8},'bold': True}})
    wb.sheet1.format(f"I4:I{row_count}", {"textFormat": {"foregroundColor": {"red": 1.0,"green": 0.43,"blue": 0.1}}})
    curr_time=datetime.datetime.now(pytz.timezone('Asia/Kolkata'))
    wb.sheet1.update([['=HYPERLINK("github.com/byethon/psms-to-gsheet","Sheet automatically updated using Github Actions + pypsd_bot(github.com/byethon/psms-to-gsheet)")']]+[[f'Sheet Last updated at {curr_time.strftime("%b %d %Y %H:%M%p")} next update at {(curr_time+datetime.timedelta(minutes=30)).strftime("%b %d %Y %H:%M%p")}']]+[dataframe.columns.values.tolist()] + dataframe.values.tolist(),value_input_option="USER_ENTERED")
    print("Program executed Successfuly")

