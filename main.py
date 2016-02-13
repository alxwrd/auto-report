import csv
import random
import sys
import os
import folium
from collections import defaultdict
from datetime import datetime
from time import gmtime, strftime

__version__ = 1.0

print "\nStarting AutoHandover - Version {}".format(__version__)
print "Created by Alex Ward"


map_osm = folium.Map(location=[52.857202, -1.298107], tiles='Stamen Toner',
                    zoom_start=7)

#check sys variables(drag and drop onto shortcut), if not present, ask for file name.
if len(sys.argv) <= 1:
      ask_sheet = raw_input('Enter csv file name: ') + '.csv'
else:
    ask_sheet = sys.argv[1]

#list of months to check first line of input file.
months = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec']

#lists of engineers so report can determine which area the report is for.
north_east = ['BLANK']
north_west = ['BLANK']
north_mids = ['BLANK']
south_mids = ['ALAN', 'CHRIS', 'DARREN', 'DAVID']
south_wales = ['BLANK']
london = ['BLANK']

#list of job codes so type of job can be determined.
repair_codes = ['MEMGRF', 'MEMGRB', 'MRMGR2', 'MEMGRC', 'MEMGR1', 'MEMGIR', \
                'MEMGR3', 'REPSER', 'SDREP', 'NCRTM', 'NPOREP', 'HCMREP', 'SCREP']
service_codes = ['MEMGIF', 'MEMGSF', 'MEMGIB', 'MEMGSB', 'MEMGIF', 'MEMGIB', \
                 'FPESN', 'NCGFSN', 'NCGFSS', 'NCWHSN', 'NCWHSD', 'NCGSIC', \
                 'NCGSID', 'NPOSER', 'HCMSER', 'SCSERV', 'MAG01']

reatnd_codes = ['REATND']
foc_codes = ['FOC']
avs_codes = ['HESSR4', 'FPSWCO', 'NCSV5', 'NCWHSO', 'NCSV3', 'FPESOW', 'FPESNW', \
             'FPSWHP', 'FPSHP', 'FPERO', 'NCSV2', 'NCREP', 'HESPR7', 'HESRP6', \
             'HESRP3', 'HESRP2', 'WRT03', 'NGSIDO', 'PETREP', 'HESRP8', 'HESRP5', \
             'HESRP4', 'HESSR9', 'POWF', 'TRV03', 'NCCBS', 'MAG02', 'NCXS25', \
             'HEESR5', 'SDPSER', 'FPERNW', 'TRV01', 'TRV02', 'FPEROW', 'WRT02', 'WRT01']

#set area and fm as global variables to be changed later.             
area = ''
fm = ''


def am_or_pm():
#ask the user when the report is for
    ask = ''
    allowed = ['1', '2', '3', '4', '5']
    print 'Choose type of report \n1) AM \n2) PM \n3) Next day\n4) Saturday\n5) Sunday\n'
    while ask not in allowed:
        ask = raw_input('Enter choice: ')
    if ask == '1':
        return 'AM'
    if ask == '2':
        return 'PM'
    if ask == '3':
        return 'Next day'
    if ask == '4':
        return 'Saturday'
    if ask == '5':
        return 'Sunday'


def current_time():
#get the current time to use as a time stamp on the report
    return strftime("%d/%m/%Y @ %H:%M:%S", gmtime())


def sheet_to_list(sheet):
#takes the csv file and adds each line as a list, to a list
    with open(sheet, 'rb') as csvfile:
        temp = csv.reader(csvfile, delimiter=',')
        final = []
        count = 0
        for row in temp:
            if len(row) == 0:
                continue
            if row[0][:3].lower() in months and row[0][9:10] == ':':
                continue
            else:
                final.append(row)
    return final


def split_engineers(sheet):
#go through the list and find engineers so we can split the shifts up.
#an engineers shift is determined by the length of the 'job' being 1
#which indicates the log on or log off.
    list_var = sheet_to_list(sheet)
    result = defaultdict(list)
    diary = False
    current = None
    name = ""
    for x in range(0,len(list_var)):
        if len(list_var[x]) == 1:
            diary = True
            current = list_var[x][0]
            result[current] = []
        if len(list_var[x]) > 1 and diary == True:
            result[current].append(list_var[x])
    return result

  
def number_of_engineers(diary_list):
#return the length of the engineer list.
    return len(diary_list)

  
def total_jobs(diary_list):
#find all jobs which is defined by jobs that have a charge code.
    result = {}
    for diary in diary_list:
        for x in range(0, len(diary_list[diary])):
            if diary_list[diary][x][2] in result:
                if diary_list[diary][x][2] != ' ' or diary_list[diary][x][2] != '':
                    result[diary_list[diary][x][2]] += 1
            elif diary_list[diary][x][2] == ' ' or diary_list[diary][x][2] == '':
                pass
            else:
                if diary_list[diary][x][2] != ' ' or diary_list[diary][x][2] != '':
                    result[diary_list[diary][x][2]] = 1
    return sum(result.values())

  
def total_travel(diary_list):
#add up all km in an engineers diary
    result = 0
    for diary in diary_list:
        for x in range(0, len(diary_list[diary])):
            result += float(diary_list[diary][x][3])
    return result

  
def get_average(total, how_many):
#get an average
    return float(total) / float(how_many)
  
diary_list = split_engineers(ask_sheet)
     
def get_name(diary_list):
#get the name of the engineer by removing the shiftid from the end.
    shift_id = ''
    result = []
    count = 0
    for diary in diary_list:
        if count == 0:
            shift_id = diary
            count += 1
        if count >= 1:
            break
    for letter in shift_id:
        if letter.isalpha() == True:
            result.append(letter)
        else:
            pass
    return ''.join(result)

  
def get_date(diary_list):
#get the date from a job.
    result = ''
    count = 0
    for diary in diary_list:
        for job in diary_list[diary]:
            if count == 0:
                result = job[6][:6]
            if count >= 1:
                break
    return result

  
def count_jobs(diary_list):
#count jobs for each engineer.
    result = {}
    for diary in diary_list:
        for x in range(0, len(diary_list[diary])):
            if diary_list[diary][x][2] in result:
                if diary_list[diary][x][2] != ' ' or diary_list[diary][x][2] != '':
                    result[diary_list[diary][x][2]] += 1
            elif diary_list[diary][x][2] == ' ' or diary_list[diary][x][2] == '':
                pass
            else:
                if diary_list[diary][x][2] != ' ' or diary_list[diary][x][2] != '':
                    result[diary_list[diary][x][2]] = 1
    return result

def sum_count_jobs():
#counts up the types of jobs and puts them into categories.
    list_var = count_jobs(diary_list)
    result = {'Repair' : 0, 'Service' : 0, 'Reattend' : 0, 'FOC' : 0, 'AVS' : 0, 'Other' : 0}
    for item in list_var:
        if item in repair_codes:
            result['Repair'] += list_var[item]
        elif item in service_codes:
            result['Service'] += list_var[item]
        elif item in reatnd_codes:
            result['Reattend'] += list_var[item]
        elif item in foc_codes:
            result['FOC'] += list_var[item]
        elif item in avs_codes:
            result['AVS'] += list_var[item]
    result['Other'] = sum(list_var.values()) - sum(result.values())
    return result


def write_job_count():
#writes the job counts as a HTML table.
    result = ''
    job_count = sum_count_jobs()
    for x in job_count:
        result += '''
        <tr><td style="width: 50%; border: 1px solid black"><b>{type}:</td>
        <td style="width: 50%; border: 1px solid black"><center>{count}</center></td></tr>'''.format({}, type=x, count=job_count[x])
    return result


def check_area(diary_list):
#checks the area backs 
    global area
    global fm
    if get_name(diary_list) in north_east:
        area = 'North East'
        fm = 'Mr Smith'
    if get_name(diary_list) in north_west:
        area = 'North West'
        fm = 'Mr Smith'
    if get_name(diary_list) in north_mids:
        area = 'North Midlands'
        fm = 'Mr Smith'
    if get_name(diary_list) in south_mids:
        area = 'South Midlands'
        fm = 'Mr Smith'
    if get_name(diary_list) in south_wales:
        area = 'South & Wales'        
        fm = 'Mr Smith'
    if get_name(diary_list) in london:
        area = 'London'
        fm = 'Mr Smith'


def readable_diary(diary_list):
    result = []
    for diary in diary_list:
        result.append({'Engineer': diary})
        job_count = 0
        for x in range(0, len(diary_list[diary])):
            if diary_list[diary][x][0][:2].isalpha() == False:
                job_count += 1
            elif diary_list[diary][x][0][:2].isalpha() == True:
                pass
        for x in range(len(result)):
            if result[x]['Engineer'] == diary: 
                result[x]['Job Count'] = job_count
            
        complete_count = 0
        for x in range(0, len(diary_list[diary])):
            if diary_list[diary][x][8] == 'COMPLETED':
                if diary_list[diary][x][0][:2].isalpha() == False:
                    complete_count += 1
            elif diary_list[diary][x][8] == ' ' or diary_list[diary][x][6] == '':
                pass
        for x in range(0, len(result)):
            if result[x]['Engineer'] == diary: 
                result[x]['Complete Count'] = complete_count
         
        km_total = 0
        for x in range(0, len(diary_list[diary])):
            km_total += float(diary_list[diary][x][3])
            for x in range(len(result)):
                if result[x]['Engineer'] == diary:
                    result[x]['Km'] = km_total
                    
        finish_time = ''
        for x in range(0, len(diary_list[diary])):
            finish_time = diary_list[diary][x][7]
            for x in range(len(result)):
                if result[x]['Engineer'] == diary:
                    result[x]['Finish Time'] = finish_time
                    
        start_time = diary_list[diary][0][7]
        for x in range(len(result)):
            if result[x]['Engineer'] == diary:
                result[x]['Start Time'] = start_time
            
    return result


def find_parts_fits(diary_list):
    result = []
    for diary in diary_list:
        for x in diary_list[diary]:
            if x[0][0].isalpha() == True:
                continue
            if int(x[4][:2]) >= 1 and int(x[0][10:11]) > 1 and (x[2] in repair_codes or x[2] in reatnd_codes):
                x.append(diary)
                result.append(x)
    return result


def find_long_duration(diary_list):
    result = []
    for diary in diary_list:
        for x in diary_list[diary]:
            if int(x[4][:2]) >= 2 and x[0][:2].isalpha() == False:
                x.append(diary)
                result.append(x)
    return result


def find_install(diary_list):
    result = []
    for diary in diary_list:
        for x in diary_list[diary]:
            if x[2] == 'INSTIN' or x[2] == 'INSTWT':
                x.append(diary)
                result.append(x)
    return result


def create_map(diary_list):
    count = 0
    for diary in diary_list:
        line_list = []
        for x in diary_list[diary]:
            if x[0][0].isalpha() == True:
                if x[9] == '0':
                    continue
                if x[4][3:5] == '15':
                    map_osm.polygon_marker([x[9], x[10]], radius=5, popup=x[0] + ' - ' + x[2], line_color='#585858', fill_color='#FAFAFA', num_sides=5, rotation=45)
                    line_list.append([float(x[9]), float(x[10])])
                else:
                    map_osm.polygon_marker([x[9], x[10]], radius=5, popup=x[0] + ' - ' + x[2], line_color='#585858', fill_color='#FAFAFA', num_sides=10, rotation=45)
                    line_list.append([float(x[9]), float(x[10])])
            elif x[2] in repair_codes:
                map_osm.polygon_marker([x[9], x[10]], radius=5, popup=x[0] + ' - ' + x[2], line_color='#5882FA', fill_color='#5882FA', num_sides=4, rotation=45)
                line_list.append([float(x[9]), float(x[10])])
            elif x[2] in service_codes:
                map_osm.polygon_marker([x[9], x[10]], radius=5, popup=x[0] + ' - ' + x[2], line_color='#F5A9F2', fill_color='#F5A9F2', num_sides=4, rotation=45)
                line_list.append([float(x[9]), float(x[10])])
            elif x[2] in avs_codes:
                map_osm.polygon_marker([x[9], x[10]], radius=5, popup=x[0] + ' - ' + x[2], line_color='#F7D358', fill_color='#F7D358', num_sides=4, rotation=45)
                line_list.append([float(x[9]), float(x[10])])
            else:
                map_osm.polygon_marker([x[9], x[10]], radius=5, popup=x[0] + ' - ' + x[2], line_color='#F78181', fill_color='#F78181', num_sides=4, rotation=45)
                line_list.append([float(x[9]), float(x[10])])
        if count == 0:
            map_osm.line(locations=line_list, line_weight=3, line_color='#08088A')
        elif count == 1:
            map_osm.line(locations=line_list, line_weight=3, line_color='#088A08')
        elif count == 2:
            map_osm.line(locations=line_list, line_weight=3, line_color='#868A08')
        elif count == 3:
            map_osm.line(locations=line_list, line_weight=3, line_color='#610B38')
        elif count >= 4:
            map_osm.line(locations=line_list, line_weight=3, line_color='#4C0B5F')
        count += 1
        if count > 4:
            count = 0
            
def find_gsi(diary_list):
    result = []
    for diary in diary_list:
        for x in diary_list[diary]:
            if x[2] == 'HESSR1':
                x.append(diary)
                result.append(x)
    return result


def find_reattend(diary_list):
    result = []
    for diary in diary_list:
        for x in diary_list[diary]:
            if x[2] == 'REATND':
                x.append(diary)
                result.append(x)
    return result
    
    
def find_multi_visit(diary_list):
    result = []
    for diary in diary_list:
        for x in diary_list[diary]:
            if x[0][:2].isalpha() == False and int(x[0][10:11]) > 2:
                x.append(diary)
                result.append(x)
    return result
    
    
def find_blocks(diary_list):
    result = []
    for diary in diary_list:
        for x in diary_list[diary]:
            if x[0][:2].isalpha() == True and (int(x[4][1:2]) >= 1 or int(x[4][3:5]) > 15):
                print x[4][1:2]
                x.append(diary)
                result.append(x)
    return result

def check_travel(x):
    if x >= 100:
        return 'bgcolor="#F78181"'
    else:
        return ' '
    # ''' + str(check_travel(int(diary['Km']))) + '''
    
    
def write_diaries():
    result = '''<table style="width: 100%; margin: 0 0 2px 2px">
    <caption><font size="2">Diary run down for ''' + area + '''</font></caption>
    <tr><td style="width: 40%; border-collapse: collapse"></td><tr></table>'''
    diaries = readable_diary(diary_list)
    for diary in diaries:
        result += '''<table style="width: 100%; border-collapse: collapse; margin: 0 0 2px 2px"><tr>
        <td style="width: 30%; border: 1px solid black; border-collapse: collapse"><font size="1.5"><b>Name:&emsp;</b> {engineer}</td>
        <td style="width: 15%; border: 1px solid black; border-collapse: collapse"><font size="1.5"><b>Jobs:&emsp;</b> {completed} / {total}</td>
        <td style="width: 20%; border: 1px solid black; border-collapse: collapse"><font size="1.5"><b>Start time:&emsp;</b> {start}</td>
        <td style="width: 20%; border: 1px solid black; border-collapse: collapse"><font size="1.5"><b>End time:&emsp;</b> {finish}</td>
        <td style="width: 15%; border: 1px solid black; border-collapse: collapse"'''.format({}, engineer=diary['Engineer'][:-11], completed=diary['Complete Count'], total=diary['Job Count'], finish=diary['Finish Time'][-5:], start=diary['Start Time'][-5:]) + str(check_travel(int(diary['Km']))) + '''><font size="1.5">{km} Km</td>
        </tr></table>
        '''.format({}, km=int(diary['Km']))
    return result


def write_reattend():
    result = '''<table style="width: 100%; margin: 0 0 2px 2px">
    <caption><font size="2">Reattends for ''' + area + '''</font></caption>
    <tr><td style="width: 40%; border-collapse: collapse"></td><tr></table>'''
    diaries = find_reattend(diary_list)
    if len(diaries) >= 1:
        for diary in diaries:
            result += '''<table style="width: 100%; border-collapse: collapse; margin: 0 0 2px 2px"><tr>
                <td style="width: 40%; border: 1px solid black; border-collapse: collapse"><font size="1.5"><b>Name:&emsp;</b> {engineer}</td>
                <td style="width: 20%; border: 1px solid black; border-collapse: collapse"><font size="1.5"><b>Call #:&emsp;</b> {job_number}</td>
                <td style="width: 20%; border: 1px solid black; border-collapse: collapse"><font size="1.5"><b>Duration:&emsp;</b> {duration}</td>
                <td style="width: 20%; border: 1px solid black; border-collapse: collapse"><font size="1.5"><b>Postcode:&emsp;</b> {postcode}</td>
                </tr></table>'''.format({}, engineer=diary[12][:-11], job_number=diary[0][:-7], duration=diary[4][:5], postcode=diary[1])
    else:
        result += '''<table style="width: 100%; border-collapse: collapse; margin: 0 0 2px 2px"><tr>
            <td style="width: 40%; border: 1px solid black; border-collapse: collapse">
            <font size="2.5"><b><center>No parts fits booked</td>
            </tr></table>
            '''
    return result


def sum_parts_fits():
    hours = 0
    minutes = 0
    diaries = find_parts_fits(diary_list)
    for diary in diaries:
        hours += int(diary[4][:2])
        minutes += int(diary[4][3:5])
    hours += int(minutes / 60)
    minutes = minutes % 60
    result = [hours, minutes, len(find_parts_fits(diary_list))]
    return result


def write_part_fits_summary():
    result = ''
    diaries = sum_parts_fits()
    result += '''<table style="width: 100%; border-collapse: collapse; margin: 0 0 2px 2px"><tr>
    <td style="width: 50%; border-collapse: collapse"><font size="2"><b>Total part fits:&emsp;</b> {total}</td>
    <td style="width: 50%; border-collapse: collapse"><font size="2"><b>Total duration:&emsp;</b> {h}:{m}</td>
    </tr></table>
    '''.format({}, total=len(find_parts_fits(diary_list)), h=diaries[0], m=diaries[1], average=())
    return result


def write_long_duration():
    result = '''<table style="width: 100%; margin: 0 0 2px 2px">
    <caption><font size="2">Long duration for ''' + area + '''</font></caption>
    <tr><td style="width: 40%; border-collapse: collapse"></td><tr></table>'''
    diaries = find_long_duration(diary_list)
    if len(diaries) > 1:
        for diary in diaries:
            result += '''<table style="width: 100%; border-collapse: collapse; margin: 0 0 2px 2px"><tr>
            <td style="width: 40%; border: 1px solid black; border-collapse: collapse"><font size="1.5"><b>Name:&emsp;</b> {engineer}</td>
            <td style="width: 20%; border: 1px solid black; border-collapse: collapse"><font size="1.5"><b>Call #:&emsp;</b> {job_number}</td>
            <td style="width: 20%; border: 1px solid black; border-collapse: collapse"><font size="1.5"><b>Duration:&emsp;</b> {duration}</td>
            <td style="width: 20%; border: 1px solid black; border-collapse: collapse"><font size="1.5"><b>Postcode:&emsp;</b> {postcode}</td>
            </tr></table>
            '''.format({}, engineer=diary[12][:-11], job_number=diary[0][:-7], duration=diary[4][:5], postcode=diary[1])
    else:
        result += '''<table style="width: 100%; border-collapse: collapse; margin: 0 0 2px 2px"><tr>
            <td style="width: 40%; border: 1px solid black; border-collapse: collapse">
            <font size="2.5"><b><center>No long duration jobs booked</td>
            </tr></table>
            '''
    return result


def write_find_install():
    result = '''<table style="width: 100%; margin: 0 0 2px 2px">
    <caption><font size="2">Install jobs for ''' + area + '''</font></caption>
    <tr><td style="width: 40%; border-collapse: collapse"></td><tr></table>'''
    diaries = find_install(diary_list)
    if len(diaries) >= 1:
        for diary in diaries:
            result += '''<table style="width: 100%; border-collapse: collapse; margin: 0 0 2px 2px"><tr>
            <td style="width: 40%; border: 1px solid black; border-collapse: collapse"><font size="1.5"><b>Name:&emsp;</b> {engineer}</td>
            <td style="width: 20%; border: 1px solid black; border-collapse: collapse"><font size="1.5"><b>Call #:&emsp;</b> {job_number}</td>
            <td style="width: 20%; border: 1px solid black; border-collapse: collapse"><font size="1.5"><b>Duration:&emsp;</b> {duration}</td>
            <td style="width: 20%; border: 1px solid black; border-collapse: collapse"><font size="1.5"><b>Postcode:&emsp;</b> {postcode}</td>
            </tr></table>
            '''.format({}, engineer=diary[12][:-11], job_number=diary[0][:-7], duration=diary[4][:5], postcode=diary[1])
    else:
        result += '''<table style="width: 100%; border-collapse: collapse; margin: 0 0 2px 2px"><tr>
            <td style="width: 40%; border: 1px solid black; border-collapse: collapse">
            <font size="2.5"><b><center>No Install jobs booked</td>
            </tr></table>
            '''
    return result


def write_find_gsi():
    result = '''<table style="width: 100%; margin: 0 0 2px 2px">
    <caption><font size="2">GSI jobs for ''' + area + '''</font></caption>
    <tr><td style="width: 40%; border-collapse: collapse"></td><tr></table>'''
    diaries = find_gsi(diary_list)
    if len(diaries) >= 1:
        for diary in diaries:
            result += '''<table style="width: 100%; border-collapse: collapse; margin: 0 0 2px 2px"><tr>
            <td style="width: 40%; border: 1px solid black; border-collapse: collapse"><font size="1.5"><b>Name:&emsp;</b> {engineer}</td>
            <td style="width: 20%; border: 1px solid black; border-collapse: collapse"><font size="1.5"><b>Call #:&emsp;</b> {job_number}</td>
            <td style="width: 20%; border: 1px solid black; border-collapse: collapse"><font size="1.5"><b>Duration:&emsp;</b> {duration}</td>
            <td style="width: 20%; border: 1px solid black; border-collapse: collapse"><font size="1.5"><b>Postcode:&emsp;</b> {postcode}</td>
            </tr></table>
            '''.format({}, engineer=diary[12][:-11], job_number=diary[0][:-7], duration=diary[4][:5], postcode=diary[1])
    else:
        result += '''<table style="width: 100%; border-collapse: collapse; margin: 0 0 2px 2px"><tr>
            <td style="width: 40%; border: 1px solid black; border-collapse: collapse">
            <font size="2.5"><b><center>No GSI jobs booked</td>
            </tr></table>
            '''
    return result


def write_find_multi_visit():
    result = '''<table style="width: 100%; margin: 0 0 2px 2px">
    <caption><font size="2">Jobs on 3rd plus visit ''' + area + '''</font></caption>
    <tr><td style="width: 40%; border-collapse: collapse"></td><tr></table>'''
    diaries = find_multi_visit(diary_list)
    if len(diaries) >= 1:
        for diary in diaries:
            result += '''<table style="width: 100%; border-collapse: collapse; margin: 0 0 2px 2px"><tr>
            <td style="width: 40%; border: 1px solid black; border-collapse: collapse"><font size="1.5"><b>Name:&emsp;</b> {engineer}</td>
            <td style="width: 20%; border: 1px solid black; border-collapse: collapse"><font size="1.5"><b>Call #:&emsp;</b> {job_number}</td>
            <td style="width: 20%; border: 1px solid black; border-collapse: collapse"><font size="1.5"><b>Duration:&emsp;</b> {duration}</td>
            <td style="width: 20%; border: 1px solid black; border-collapse: collapse"><font size="1.5"><b>Postcode:&emsp;</b> {postcode}</td>
            </tr></table>
            '''.format({}, engineer=diary[12][:-11], job_number=diary[0][:-3], duration=diary[4][:5], postcode=diary[1])
    else:
        result += '''<table style="width: 100%; border-collapse: collapse; margin: 0 0 2px 2px"><tr>
            <td style="width: 40%; border: 1px solid black; border-collapse: collapse">
            <font size="2.5"><b><center>No jobs on 3rd plus visit</td>
            </tr></table>
            '''
    return result

    
def write_find_blocks():
    result = '''<table style="width: 100%; margin: 0 0 2px 2px">
    <caption><font size="2">Blocks booked for ''' + area + '''</font></caption>
    <tr><td style="width: 40%; border-collapse: collapse"></td><tr></table>'''
    diaries = find_blocks(diary_list)
    if len(diaries) >= 1:
        for diary in diaries:
            result += '''<table style="width: 100%; border-collapse: collapse; margin: 0 0 2px 2px"><tr>
            <td style="width: 40%; border: 1px solid black; border-collapse: collapse"><font size="1.5"><b>Name:&emsp;</b> {engineer}</td>
            <td style="width: 20%; border: 1px solid black; border-collapse: collapse"><font size="1.5"><b>Block ID:&emsp;</b> {job_number}</td>
            <td style="width: 20%; border: 1px solid black; border-collapse: collapse"><font size="1.5"><b>Duration:&emsp;</b> {duration}</td>
            <td style="width: 20%; border: 1px solid black; border-collapse: collapse"><font size="1.5"><b>Postcode:&emsp;</b> {postcode}</td>
            </tr></table>
            '''.format({}, engineer=diary[12][:-11], job_number=diary[0][:-4], duration=diary[4][:5], postcode=diary[1])
    else:
        result += '''<table style="width: 100%; border-collapse: collapse; margin: 0 0 2px 2px"><tr>
            <td style="width: 40%; border: 1px solid black; border-collapse: collapse">
            <font size="2.5"><b><center>No blocks booked</td>
            </tr></table>
            '''
    return result   
    
ampm = am_or_pm()

def create_html_header():
    check_area(diary_list)
    with open((area + 'report' + get_date(diary_list)+ '.htm'), 'w') as report:
        report.write('''
        <title>
        ''' + area + ''' Handover
        </title>
        <font face="Calibri">

        <b><font size="6">''' + area + ''' <i>''' + ampm + '''</i> Handover</font></b> for <i>''' + get_date(diary_list) + '''</i><br>
        <font size="5">''' + fm + '''<br>
        <br><font size="1"><i>&emsp;Generated: ''' + current_time() + '''</i></font>
        <br><br>
        <table style="width:100%; border-collapse: collapse; margin: 10 10 10px 0px">
        <caption><font size="3"><b>Figures for the ''' + area + '''</b></font></caption>
        <tr><td style="width: 50%; border: 1px solid black"><b># of engineers: </td>
        <td style="width: 50%; border: 1px solid black"><center>''' + str(number_of_engineers(diary_list)) + '''</center> </td></tr>

        <tr><td style="border: 1px solid black; border-collapse: collapse"><b>Total jobs: </td>
        <td style="border: 1px solid black"><center>''' + str(total_jobs(diary_list)) + ''' </center></td></tr>

        <tr><td style="border: 1px solid black; border-collapse: collapse"><b>Total travel: </td>
        <td style="border: 1px solid black; border-collapse: collapse"><center>''' + "{0:.2f}".format(total_travel(diary_list)) + ''' </center></td></tr>

        <tr><td style="border: 1px solid black; border-collapse: collapse"><b>Average jobs: </td>
        <td style="border: 1px solid black; border-collapse: collapse"><center>''' + "{0:.2f}".format(get_average(total_jobs(diary_list), number_of_engineers(diary_list))) + ''' </center></td></tr>
        
        <tr><td style="border: 1px solid black; border-collapse: collapse"><b>Average travel: </td>
        <td style="border: 1px solid black; border-collapse: collapse"><center>''' + "{0:.2f}".format(get_average(total_travel(diary_list), number_of_engineers(diary_list))) + '''</center> </td></tr>
        
        <tr><td style="border: 1px solid black; border-collapse: collapse"><b>Parts fits: </td>
        <td style="border: 1px solid black; border-collapse: collapse"><center>''' + str(len(find_parts_fits(diary_list))) + '''</center> </td></tr>
        
        <tr><td style="border: 1px solid black; border-collapse: collapse"><b>Parts duration: </td>
        <td style="border: 1px solid black; border-collapse: collapse"><center>''' + str(sum_parts_fits()[0]) + ''':''' + str(sum_parts_fits()[1]) + '''</center> </td></tr></table>''')
        report.write('<table style="width:100%; border-collapse: collapse; margin: 10 10 10px 0px">')
        report.write('''<caption><font size="3"><b>Job breakdown for ''' + area + ''' </b></font></caption>''')
        report.write(write_job_count())
        report.write('</table>')
        
def create_html_body():
    job_count = count_jobs(diary_list)
    with open((area + 'report' + get_date(diary_list)+ '.htm'), 'a') as report:
        report.write('<br><hr align="left" width="100%">')      
        report.write(write_diaries())
        report.write('<br><hr align="left" width="100%">')
        report.write(write_long_duration())
        report.write('<br><hr align="left" width="100%">')
        report.write(write_reattend())
        report.write('<br><hr align="left" width="100%">')
        report.write(write_find_multi_visit())
        report.write('<br><hr align="left" width="100%">')
        report.write(write_find_blocks())
        report.write('<br><hr align="left" width="100%">')
        report.write(write_find_install())
        report.write('<br><hr align="left" width="100%">')
        report.write(write_find_gsi())

        
def create_html_footer():
    with open((area + 'report' + get_date(diary_list)+ '.htm'), 'a') as report:
        report.write('''
        <table style="width:100%; border-collapse: collapse; margin: 10 10 10px 0px"><tr><td><font size="2">
        <center><i>Automatically generated with data from ORS.
        All information was complete at time of creation.</i>
        </td></tr></table>
        ''')


def create_report():
    create_html_header()
    create_html_body()
    create_html_footer()

create_report()
create_map(diary_list)
map_osm.create_map(path='{}Map{}.htm'.format(area, get_date(diary_list)))


print "Creating report for %s... \n " % area
print "# of engineers:\t\t %s" % number_of_engineers(diary_list)
print "Total jobs:\t\t %s" % total_jobs(diary_list)
print "Total travel:\t\t %.2f" % total_travel(diary_list)
print "Average jobs:\t\t %.2f" % get_average(total_jobs(diary_list), number_of_engineers(diary_list))
print "Average travel:\t\t %.2f" % get_average(total_travel(diary_list), number_of_engineers(diary_list))
