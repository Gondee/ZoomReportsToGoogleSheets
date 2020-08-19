import json
from zoomus import ZoomClient
import pymongo
import gspread
import time
from datetime import datetime, date, timedelta

#LoopingSpeed
SLEEP_DURATION = 60 * 60 # 60 seconds x 60

#MongoDB cloud information
MONGO_CONNECTION_STRING = ""
DB_MEETING_NAME = "Reports" #DB name
DB_COLLECTION_MEETINGS = 'OrgOverview' #collection for meeting list
DB_COLLECTION_PARTICIPATION = 'meetings' #collection for individual meeting list and caching.

#Zoom api information
ZOOM_API_KEY = ""
ZOOM_API_PASS = ''

#admin email for google sheets
ADMIN_EMAIL = ''
GSUIT_SERVICE_ACCOUNT_PATH = ''

#Show debug statements of nio?
SUPRESS_DEBUG_PRINT = True


def debug_print(printMe):
    if(SUPRESS_DEBUG_PRINT == False):
        print(printMe)

#time.sleep(200) #Delay for starting

while(True):
    gc = gspread.service_account(filename=GSUIT_SERVICE_ACCOUNT_PATH)
    pyMongoClient = pymongo.MongoClient(MONGO_CONNECTION_STRING)
    db = pyMongoClient[DB_MEETING_NAME]
    masterReport = db[DB_COLLECTION_MEETINGS]
    slaveReport = db[DB_COLLECTION_PARTICIPATION]
    masterReport.create_index([('uuid', pymongo.ASCENDING),('start_time', pymongo.ASCENDING)], unique=True) #Remove ability to insert duplicates to skirt time issues
    slaveReport.create_index([('uuid', pymongo.ASCENDING),('name', pymongo.ASCENDING)], unique=True) #Remove ability to insert duplicates to skirt time issues
    client = ZoomClient(ZOOM_API_KEY, ZOOM_API_PASS, version=2)
    user_list_response = client.user.list()
    user_list = json.loads(user_list_response.content)

    #Caching non duplicates to push to google sheets
    meetingsCache = []
    participantsCache = []

    try:
        sh = gc.open('ZoomDashboard')
    except:
        debug_print("No sheet, creating and sharing")
        sh = gc.create('ZoomDashboard')
        sh.share(ADMIN_EMAIL, perm_type='user', role='writer')
        bh = gc.open('ZoomDashboard').get_worksheet(1)  # get the worksheet for meetings.
        shh = gc.open('ZoomDashboard').sheet1  # open First page
        bh.resize(1)
        shh.resize(1)


    sh = gc.open('ZoomDashboard').sheet1 #open First page
    bh = gc.open('ZoomDashboard').get_worksheet(1) #get the worksheet for meetings.
    bh.resize(1)
    sh.resize(1)
    for user in user_list['users']:
        user_id = user['id']
        start_time = datetime.now() - timedelta(seconds=SLEEP_DURATION)
        end_time = datetime.utcnow()

        #take each meeting host, and see what meetings they been in.
        userReport = client.report.get_user_report(user_id=user_id, start_time=start_time, end_time=end_time) #Parse JSON
        meetingsList = (json.loads(userReport.content))['meetings'] #pull out the list of meetings


        #Insert the meetings into the DB
        for meeting in meetingsList:
            try:
                #debug_print(str('Attempting Insert - ' + meeting['id']))
                masterReport.insert_one(meeting)
                meetingsCache.append(meeting)
            except:
                debug_print('Error Inserting a Doc - Probably a Duplicate')




        #get particpiants doesnt seem to get them, this could be because of the buisness plan or because the api wrapper is messed up.
        for meeting in meetingsList:
            m = client.report.get_meeting_participant_report(meeting_id=str(meeting['uuid']))
            participants = json.loads(m.content)['participants']

            #check to ensure there is actually particpants
            if(len(participants)!=0):
                #Add the meeting_id to each of these participant records.
                for p in participants:
                    try:
                        p['meeting_id'] = meeting['id']
                        p['uuid'] = meeting['uuid']
                        p['start_time'] = meeting['start_time']
                        slaveReport.insert_one(p)
                        participantsCache.append(p)
                    except:
                        debug_print('Duplicate participant entry')
            else:
                debug_print('There is no participants of this meeting so not adding')


    #Add the new Meetings to the Sheet
    rows = []
    for m in meetingsCache:
        row = [m['id'],m['host_id'],m['uuid'],m['type'],m['topic'],m['user_name'],m['user_email'],m['start_time'],m['end_time'],m['duration'],m['total_minutes'],m['participants_count']]
        rows.append(row)
    bh.append_rows(rows)

    parts = []
    #Add the new particpants list
    for p in participantsCache:
        row = [p['id'], p['name'], p['user_email'], p['meeting_id'],p['uuid'],p['start_time']]
        parts.append(row)
    sh.append_rows(parts)

    pyMongoClient.close() #CLose for reconnection later
    print("Sleeping")
    time.sleep(SLEEP_DURATION-1)
