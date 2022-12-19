import pandas as pd
import time
import os
import numpy as np
from openpyxl import Workbook, load_workbook
import streamlit as st
from io import BytesIO


st.set_page_config(page_title='Test')
hide_streamlit_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            </style>
            """
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

class TimeString_HHMM():
    def __init__(self, input_string = None):
        #split at position 0 up to first colon and fill 2 0s - HOURS 
        self.hh = input_string.split(':')[0].zfill(2)
        #split at position 1 up to first colon and fill 2 0s - MINUTES
        self.mm = input_string.split(':')[1].zfill(2)
        #Convert hours to minutes to get total minutes value by hour-minutes + minutes
        self.in_minutes = int(self.hh)*60 + int(self.mm)

class FullSchedule():
    def __init__(self, file_path = None):
        #filepath
        
        self.path = file_path

        #Create a DataFrame out of the raw Full Schedule
        df = pd.read_excel(self.path, dtype=str)
        #replace blanks with nans in event type column
        df['Event Type'].replace('', np.nan, inplace=True)
        #drop na's in event type
        df.dropna(subset=['Event Type'], inplace=True)

        self.dataFrame = df
        #create a list of service group days by the service groupe days column in the list (UNIQUE BASED ON SET)
        self.serviceGroupDaysList = list(set(self.dataFrame['Service Group Days'].to_list()))
        #create event type list by getting event type (UNIQUE BASED ON SET)
        self.eventTypeList = list(set(self.dataFrame['Event Type'].to_list()))
        #create a list from preference groups (UNIQUE BASED ON SET)
        self.prefGroups = list(set(self.dataFrame['Pref Group'].to_list()))


    #Adapt the DataFrame the way the client did in Excel 
    def insert_extra_columns(self):    
        df = self.dataFrame
        #create new columns for the dataframe using empty lists
        _TimeColumn = []
        _MeasureColumn = []
        _LayoverJoinUpColumn = []
        _PaidColumn = []
        _BreakCountColumn = []

        #Mapping for Break Count
        #creating a mapping dataframe where copying standby event type only
        mappingDf = df[df['Event Type'] == 'standby'].copy()
        #Creating a new column, concatentating duty ID and mapping df with underscore
        mappingDf['mappingColumn'] = mappingDf['Duty id'] + '_' + mappingDf['Service Group Days']
        _mappingList = mappingDf['mappingColumn'].to_list()

        for index, row in df.iterrows():
            startTime = TimeString_HHMM(row['Start Time'])
            endTime = TimeString_HHMM(row['End Time'])
            endHour = endTime
            
            if int(startTime.hh) > int(endTime.hh):
                endHour = TimeString_HHMM(f"{int(endTime.hh) + 24}:{endTime.mm}")
            
            duration = (endHour.in_minutes - startTime.in_minutes)*24*0.000694 #This is the factor that Excel uses for the Time column
            _TimeColumn.append(duration)

            if row['Event Type'] == 'standby' and (row['Description'] == 'Break' or row['Description'] == 'Paid Break'):

                _matchingCode = f"{row['Duty id']}_{row['Service Group Days']}"
                _breakCount = _mappingList.count(_matchingCode)

                _timePaidTime = duration
                if duration < 0.75:
                    _paidTimeValue = 0.00
                else:
                    _paidTimeValue = duration - 0.75

            else:
                _paidTimeValue = 0.00
                _breakCount = 0
            
            _PaidColumn.append(_paidTimeValue)
            _BreakCountColumn.append(_breakCount)


            if index == 0:
                _MeasureColumn.append('JOINUP')
                _LayoverJoinUpColumn.append(0.00)

            else:
                currentRow_eventType = row['Event Type']
                previousRow_eventType = df.iloc[index-1]['Event Type']

                if currentRow_eventType == previousRow_eventType or (currentRow_eventType == 'service_trip' and previousRow_eventType == 'depot_pull_out') or (currentRow_eventType == 'service_trip' and previousRow_eventType == 'deadhead'):
                    _MeasureColumn.append('Layover')
                else:
                    _MeasureColumn.append('JOINUP')
                
                currentRow_dutyId = row['Duty id']
                previousRow_dutyId = df.iloc[index-1]['Duty id']

                currentRow_startTime = TimeString_HHMM(row['Start Time'])
                previousRow_endTime = TimeString_HHMM(df.iloc[index-1]['End Time'])
                _startHour = currentRow_startTime

                if int(_startHour.hh) < int(previousRow_endTime.hh):
                    #Surely this should now be adding the value to previous row end time but current row start hour? 
                    #_startHour = TimeString_HHMM(f"{int(previousRow_endTime.hh) + 24}:{previousRow_endTime.mm}")

                    #THIS HAS GOT THE MAX VALUE DISTRIBUTION DOWN FOR THE LAYOVER VALUES AFTER MIDNIGHT
                    _startHour = TimeString_HHMM(f"{int(currentRow_startTime.hh) + 24}:{currentRow_startTime.mm}")
                
                if currentRow_dutyId == previousRow_dutyId:
                    _value = (_startHour.in_minutes - previousRow_endTime.in_minutes)*24*0.000694
                    if _value < 0:
                        _value = 0.00
                else:
                    _value = 0.00
                
                _LayoverJoinUpColumn.append(_value)

                #_LayoverJoinUpColumn = [round(item, 2) for item in _LayoverJoinUpColumn]

        df['Time'] = _TimeColumn
        df['Measure'] = _MeasureColumn
        df['Layover / Join Up Value'] = _LayoverJoinUpColumn
        df['Paid'] = _PaidColumn
        df['Break Count'] = _BreakCountColumn

        df['matchingKeyBreakCount'] = df['Duty id'] + '_' + df['Service Group Days']

        _auxDf = df.copy()
        _auxDf = _auxDf[_auxDf['Event Type'] == 'standby'].copy()

        _auxDf = _auxDf.sort_values('Time', ascending=False).drop_duplicates('matchingKeyBreakCount').sort_index()

        _auxDf.set_index('matchingKeyBreakCount', inplace=True)
        _BreaksCountDict = _auxDf.to_dict('index')

        _newPaidColumn = [] #Storing the new Paid values in here following the client's rules

        for index, row in df.iterrows():
            if int(row['Break Count']) <= 1:
                _newPaidColumn.append(row['Paid'])
            else:
                if row['Start Time'] == _BreaksCountDict[row['matchingKeyBreakCount']]['Start Time']:
                    _newPaidColumn.append(row['Paid'])
                else:
                    _newPaidColumn.append(row['Time'])

        df['newPaid'] = _newPaidColumn

        self.adaptadedDataFrame = df

  
        #dfpaid = df[df['Event Type'] == 'standby'].copy()
        
    
        buffer = BytesIO()
        with pd.ExcelWriter(buffer) as writer:
            df.to_excel(writer)
        st.download_button('Download Calculated Dataframe', data = buffer, file_name='str.xlsx')
        
    

    

        


        
    ########################################################################################################################################

        #CREATE A FILTERED DATAFRAME TO LOOK FOR DISCEPENCIES

  


            #MAX VALUE OF JOINUP BEING SCEWED DUE TO DEADHEAD OVERLAPPING WITH TRIP TIME: 
            #07:40 - 08:04 (start time, end time) - 5452 idx - Deadhead 
            #07:40 - 08:49 (start time, end time) - 5452 idx - Trip

            #7012:7013

            #Client has manually overidden the formula for these overlapping deadheads

       



    def build_printable_table(self, pref_group = None):
        if pref_group == None:
            pref_group = self.prefGroups
        elif type(pref_group) == str:
            pref_group = [pref_group]
        
        #THESE STRINGS WEREN"T SEPERATED SO SEPERATED TO CALCULATE ACCORDINGLY
        _signList = ['RF', 'RH', 'SHTL']
        
        df = self.adaptadedDataFrame
        df = df[df['Pref Group'].isin(pref_group)] #Here we already filter for the desired preference groups
        timePerEvetTypeDict = get_time_per_event_type_dict(df, self.eventTypeList, self.serviceGroupDaysList)
    
        paidTimeDict = get_pivot_sheet2(df, self.serviceGroupDaysList)
        layoverJoinUpDict = get_pivot_sheet5(df, self.serviceGroupDaysList)
        reliefTimeDict = get_sum_of_time_relief_cars(df, self.serviceGroupDaysList, _signList)

    
        
        #Build a new Dict for the final table plotting
        finalTableDict = {}
        finalTableDict['Pref Group'] = pref_group

        #In service Time
        #download calculated dataframe
       

    

        

       

        #WHY IS relieftime dict only getting 0's - need to check this function is correct
      

        
        
        finalTableDict['In service Time'] = {}
        finalTableDict['In service Time']['M-F'] = _monFri = round(timePerEvetTypeDict['service_trip']['256'] - reliefTimeDict['256'], 2)
        finalTableDict['In service Time']['Sat'] = _sat = round(timePerEvetTypeDict['service_trip']['7'] - reliefTimeDict['7'],2)
        finalTableDict['In service Time']['Sun'] = _sun = round(timePerEvetTypeDict['service_trip']['1'] - reliefTimeDict['1'],2)
        finalTableDict['In service Time']['Standard week'] = round(_monFri * 5 + _sat + _sun,2)
        #finalTableDict['In service Time']['Comments'] = 'In service driving time'

        

        

        #Deadtime
        finalTableDict['Deadtime'] = {}
        finalTableDict['Deadtime']['M-F'] = _monFri = round(timePerEvetTypeDict['deadhead']['256'] + timePerEvetTypeDict['depot_pull_in']['256'] + timePerEvetTypeDict['depot_pull_out']['256'],2)
        finalTableDict['Deadtime']['Sat'] = _sat = round(timePerEvetTypeDict['deadhead']['7'] + timePerEvetTypeDict['depot_pull_in']['7'] + timePerEvetTypeDict['depot_pull_out']['7'],2)
        finalTableDict['Deadtime']['Sun'] = _sun = round(timePerEvetTypeDict['deadhead']['1'] + timePerEvetTypeDict['depot_pull_in']['1'] + timePerEvetTypeDict['depot_pull_out']['1'],2)
        finalTableDict['Deadtime']['Standard week'] = round(_monFri * 5 + _sat + _sun,2)
        #finalTableDict['Deadtime']['Comments'] = 'Positioning trips time'

        #Attendance recovery time
        finalTableDict['Attendance'] = {}
        finalTableDict['Attendance']['M-F'] = _monFri = round(timePerEvetTypeDict['attendance']['256'] + layoverJoinUpDict['Layover']['256'],2) #Needs to be summed to a discrepancy that is unknown
        finalTableDict['Attendance']['Sat'] = _sat = round(timePerEvetTypeDict['attendance']['7'] + layoverJoinUpDict['Layover']['7'],2) #Needs to be summed to a discrepancy that is unknown
        finalTableDict['Attendance']['Sun'] = _sun = round(timePerEvetTypeDict['attendance']['1'] + layoverJoinUpDict['Layover']['1'],2) #Needs to be summed to a discrepancy that is unknown
        finalTableDict['Attendance']['Standard week'] = round(_monFri * 5 + _sat + _sun,2)
        #finalTableDict['Attendance']['Comments'] = 'Recovery time whilst a driver is onboard'

        #Scheduled Bus Hours
        finalTableDict['Scheduled Bus Hours'] = {}
        finalTableDict['Scheduled Bus Hours']['M-F'] = _monFri = round(finalTableDict['In service Time']['M-F'] + finalTableDict['Deadtime']['M-F'] + finalTableDict['Attendance']['M-F'],2)
        finalTableDict['Scheduled Bus Hours']['Sat'] = _sat = round(finalTableDict['In service Time']['Sat'] + finalTableDict['Deadtime']['Sat'] + finalTableDict['Attendance']['Sat'],2)
        finalTableDict['Scheduled Bus Hours']['Sun'] = _sun = round(finalTableDict['In service Time']['Sun'] + finalTableDict['Deadtime']['Sun'] + finalTableDict['Attendance']['Sun'],2)
        finalTableDict['Scheduled Bus Hours']['Standard week'] = round(_monFri * 5 + _sat + _sun,2)
        #finalTableDict['Scheduled Bus Hours']['Comments'] = 'Total time drivers are in control of a vehicle'

        #Sign on
        finalTableDict['Sign on'] = {}
        finalTableDict['Sign on']['M-F'] = _monFri = round(timePerEvetTypeDict['sign_on']['256'],2)
        finalTableDict['Sign on']['Sat'] = _sat = round(timePerEvetTypeDict['sign_on']['7'],2)
        finalTableDict['Sign on']['Sun']= _sun = round(timePerEvetTypeDict['sign_on']['1'],2)
        finalTableDict['Sign on']['Standard week'] = round(_monFri * 5 + _sat + _sun,2)
        #finalTableDict['Sign on']['Comments'] = 'Paid signing on time at start of duties'

        #Vehicle prep
        finalTableDict['Vehicle prep'] = {}
        finalTableDict['Vehicle prep']['M-F'] = _monFri = round(timePerEvetTypeDict['pre_trip']['256'],2)
        finalTableDict['Vehicle prep']['Sat'] = _sat = round(timePerEvetTypeDict['pre_trip']['7'],2)
        finalTableDict['Vehicle prep']['Sun']= _sun = round(timePerEvetTypeDict['pre_trip']['1'],2)
        finalTableDict['Vehicle prep']['Standard week'] = round(_monFri * 5 + _sat + _sun,2)
        #finalTableDict['Vehicle prep']['Comments'] = 'Vehicle first user check before leaving depot'

        #Travel time
        finalTableDict['Travel time'] = {}
        finalTableDict['Travel time']['M-F'] = _monFri = round(timePerEvetTypeDict['public_travel']['256'] + timePerEvetTypeDict['walk']['256'] + timePerEvetTypeDict['relief_car']['256'] + reliefTimeDict['256'],2)
        finalTableDict['Travel time']['Sat'] = _sat = round(timePerEvetTypeDict['public_travel']['7'] + timePerEvetTypeDict['walk']['7'] + timePerEvetTypeDict['relief_car']['7'] + reliefTimeDict['7'],2)
        finalTableDict['Travel time']['Sun']= _sun = round(timePerEvetTypeDict['public_travel']['1'] + timePerEvetTypeDict['walk']['1'] + timePerEvetTypeDict['relief_car']['1'] + reliefTimeDict['1'],2)
        finalTableDict['Travel time']['Standard week'] = round(_monFri * 5 + _sat + _sun,2)
        #finalTableDict['Travel time']['Comments'] = 'Paid travel time for drivers'

        #Join Up
        finalTableDict['Join Up'] = {}
        finalTableDict['Join Up']['M-F'] = _monFri = round(layoverJoinUpDict['JOINUP']['256'],2) #Needs to be summed to a discrepancy that is unknown
        finalTableDict['Join Up']['Sat'] = _sat = round(layoverJoinUpDict['JOINUP']['7'],2) #Needs to be summed to a discrepancy that is unknown
        finalTableDict['Join Up']['Sun'] = _sun = round(layoverJoinUpDict['JOINUP']['1'],2) #Needs to be summed to a discrepancy that is unknown
        finalTableDict['Join Up']['Standard week'] = round(_monFri * 5 + _sat + _sun,2)
        #finalTableDict['Join Up']['Comments'] = 'Paid time to join up work pieces'

        #Paid Breaks
        finalTableDict['Paid Breaks'] = {}
        finalTableDict['Paid Breaks']['M-F'] = _monFri = round(paidTimeDict['256'],2)
        finalTableDict['Paid Breaks']['Sat'] = _sat = round(paidTimeDict['7'],2)
        finalTableDict['Paid Breaks']['Sun'] = _sun = round(paidTimeDict['1'],2)
        finalTableDict['Paid Breaks']['Standard week'] = round(_monFri * 5 + _sat + _sun,2)
        #finalTableDict['Paid Breaks']['Comments'] = 'Paid element of meal breaks over 45 mins and 2nd breaks (fully paid)'

        #CHANGED depot_pull_out to post_trip 
        #Vehicle Park
        finalTableDict['Vehicle Park'] = {}
        finalTableDict['Vehicle Park']['M-F'] = _monFri = round(timePerEvetTypeDict['post_trip']['256'],2)
        finalTableDict['Vehicle Park']['Sat'] = _sat = round(timePerEvetTypeDict['post_trip']['7'],2)
        finalTableDict['Vehicle Park']['Sun']= _sun = round(timePerEvetTypeDict['post_trip']['1'],2)
        finalTableDict['Vehicle Park']['Standard week'] = round(_monFri * 5 + _sat + _sun,2)
        #finalTableDict['Vehicle Park']['Comments'] = 'Parking of bus when returning to depot'

       
        #Changed pre_trip to Sign_off
        #Sign off
        finalTableDict['Sign off'] = {}
        finalTableDict['Sign off']['M-F'] = _monFri = round(timePerEvetTypeDict['sign_off']['256'],2)
        finalTableDict['Sign off']['Sat'] = _sat = round(timePerEvetTypeDict['sign_off']['7'],2)
        finalTableDict['Sign off']['Sun']= _sun = round(timePerEvetTypeDict['sign_off']['1'],2)
        finalTableDict['Sign off']['Standard week'] = round(_monFri * 5 + _sat + _sun,2)
        #finalTableDict['Sign off']['Comments'] = 'Paying in time at end of shift'

        #Scheduled Paid Hours
        finalTableDict['Scheduled Paid Hours'] = {}
        _monFri = _sat = _sun = 0

        _relevantListSchedulePaidHours = list(set(list(finalTableDict.keys())) - set(['Scheduled Paid Hours', 'Scheduled Bus Hours', 'Pref Group']))

        for key in _relevantListSchedulePaidHours:
            _monFri += finalTableDict[key]['M-F']
            _sat += finalTableDict[key]['Sat']
            _sun += finalTableDict[key]['Sun']
        

        finalTableDict['Scheduled Paid Hours']['M-F'] = round(_monFri,2)
        finalTableDict['Scheduled Paid Hours']['Sat'] = round(_sat,2)
        finalTableDict['Scheduled Paid Hours']['Sun'] = round(_sun,2)
        finalTableDict['Scheduled Paid Hours']['Standard week'] = round(_monFri * 5 + _sat + _sun,2)
        #finalTableDict['Scheduled Paid Hours']['Comments'] = 'Total Scheduled paid time'

        #Scheduled Efficiency
        finalTableDict['Scheduled Efficiency'] = {}
        finalTableDict['Scheduled Efficiency']['M-F'] = round(finalTableDict['Scheduled Bus Hours']['M-F'] / finalTableDict['Scheduled Paid Hours']['M-F'], 3)*100
        finalTableDict['Scheduled Efficiency']['Sat'] = round(finalTableDict['Scheduled Bus Hours']['Sat'] / finalTableDict['Scheduled Paid Hours']['Sat'],3)*100
        finalTableDict['Scheduled Efficiency']['Sun'] = round(finalTableDict['Scheduled Bus Hours']['Sun'] / finalTableDict['Scheduled Paid Hours']['Sun'],3)*100
        finalTableDict['Scheduled Efficiency']['Standard week'] = round(finalTableDict['Scheduled Bus Hours']['Standard week'] / finalTableDict['Scheduled Paid Hours']['Standard week'],3)*100
        #finalTableDict['Scheduled Efficiency']['Comments'] = '% of Paid time drivers are in control of a vehicle'

        
        #Percentage of Pay
       
        _relevantListPercentagePay = list(set(list(finalTableDict.keys())) - set(['Pref Group']))
        
        for key in _relevantListPercentagePay:
            if key in ['Scheduled Paid Hours', 'Scheduled Efficiency']:
                finalTableDict[key]['% of Pay'] = ''
            else:
                finalTableDict[key]['% of Pay'] = round(finalTableDict[key]['Standard week'] / finalTableDict['Scheduled Paid Hours']['Standard week'],3) *100

        #ASSIGNED COMMENTS AFTRER % OF PAY. Not the best solution but works...
        
        finalTableDict['In service Time']['Comments'] = 'In service driving time'
        finalTableDict['Deadtime']['Comments'] = 'Positioning trips time'
        finalTableDict['Attendance']['Comments'] = 'Recovery time whilst a driver is onboard'
        finalTableDict['Scheduled Bus Hours']['Comments'] = 'Total time drivers are in control of a vehicle'
        finalTableDict['Sign on']['Comments'] = 'Paid signing on time at start of duties'
        finalTableDict['Vehicle prep']['Comments'] = 'Vehicle first user check before leaving depot'
        finalTableDict['Travel time']['Comments'] = 'Paid travel time for drivers'
        finalTableDict['Join Up']['Comments'] = 'Paid time to join up work pieces'
        finalTableDict['Paid Breaks']['Comments'] = 'Paid element of meal breaks over 45 mins and 2nd breaks (fully paid)'
        finalTableDict['Vehicle Park']['Comments'] = 'Parking of bus when returning to depot'
        finalTableDict['Sign off']['Comments'] = 'Paying in time at end of shift'
        finalTableDict['Scheduled Paid Hours']['Comments'] = 'Total Scheduled paid time'
        finalTableDict['Scheduled Efficiency']['Comments'] = '% of Paid time drivers are in control of a vehicle'
        

        self.finalTableDict = finalTableDict

#get Pivot Table for Time per Event Type
def get_time_per_event_type_dict(data_frame, eventTypeList, serviceGroupDaysList):
    df = data_frame
    time_per_event_type_dict = {}

    for event_type in eventTypeList:
        time_per_event_type_dict[event_type] = {}
        for serviceGroup in serviceGroupDaysList:
            time_per_event_type_dict[event_type][serviceGroup] = df.loc[(df['Service Group Days'] == serviceGroup) & (df['Event Type'] == event_type), 'Time'].sum()

    timePerEvetTypeDict = time_per_event_type_dict

    return timePerEvetTypeDict
    

#get Sheet 3 (aka Sum of Paid Time)
def get_pivot_sheet2(data_frame, serviceGroupDaysList):

    df = data_frame
    paidTimeDict = {}


    for serviceGroup in serviceGroupDaysList:
        paidTimeDict[serviceGroup] = df.loc[df['Service Group Days'] == serviceGroup, 'newPaid'].sum()

    return paidTimeDict

#get Sheet 5 (aka Sum of Layover / Join Up Value)
def get_pivot_sheet5(data_frame, serviceGroupDaysList):
    df = data_frame

    labelsList = list(set(df['Measure'].to_list()))
    layoverJoinUpDict = {}

    for measure in labelsList:
        layoverJoinUpDict[measure] = {}
        for serviceGroup in serviceGroupDaysList:
            layoverJoinUpDict[measure][serviceGroup] = df.loc[(df['Service Group Days'] == serviceGroup) & (df['Measure'] == measure), 'Layover / Join Up Value'].sum()
    
    return layoverJoinUpDict

#get Pivot Table Sum of Time per Relief Cars
def get_sum_of_time_relief_cars(data_frame, serviceGroupDaysList, _signList):

    _undesiredEventList = ['changeover', 'standby', 'split']

    df = data_frame
    df = df[~df['Event Type'].isin(_undesiredEventList)]
    df = df[df['Sign'].isin(_signList)]

    reliefTimeDict = {}

    for serviceGroup in serviceGroupDaysList:
        reliefTimeDict[serviceGroup] =  df.loc[df['Service Group Days'] == serviceGroup, 'Time'].sum()

    #Only first set in dictionary needs to be used
    
    return reliefTimeDict

    

#reliefTimeDict = get_sum_of_time_relief_cars(df, self.serviceGroupDaysList, _signList)

class TableBuilder():
    def __init__(self, file_path = None):
        self.file_path = file_path
        self.fs = FullSchedule(self.file_path)
        self.fs.insert_extra_columns()

        

    def buildTableFile(self):
        prefGroupList = list(set(self.fs.adaptadedDataFrame['Pref Group'].to_list()))

        #print(prefGroupList)

        newPrefGroupList = [None]

        for item in prefGroupList:
            newPrefGroupList.append(item)

        #print(newPrefGroupList)
        
        tablesList = []

        timestamp = time.strftime("%Y%m%d%H%M%S")
        
        
        _fileName = "{}_Tables.xlsx".format(timestamp)

        

        #outputPath = os.path.join(_dirName,_fileName)
        output = BytesIO()
        wb = Workbook()
        #wb.save(output)

        
        #writer = pandas.ExcelWriter(output,engine='xlsxwriter')
        #df.to_excel(writer)
        #writer.save()
        #output.seek(0)
        #workbook = output.read()

        #wb = load_workbook(output)

        ws = wb.active

        for prefGroup in newPrefGroupList:
            self.fs.build_printable_table(prefGroup)
            newTable = self.fs.finalTableDict
            
            tablesList.append(newTable)
        
        self.tablesList = tablesList

        #print(tablesList)

        _headerList = ['Pay Category', 'M-F', 'Sat', 'Sun', 'Standard week', '% of Pay', 'Comments']
        _subKeyList = ['M-F', 'Sat', 'Sun', 'Standard Week', '% of Pay', 'Comments']

        for index, _table in enumerate(self.tablesList):
            current_dict = self.tablesList[index]

            ws.append(['Summary of Paid Driver hours in a normal week'])
            ws.append(current_dict['Pref Group'])
            ws.append([''])
            ws.append(_headerList)

            for key in current_dict.keys():
                if key != 'Pref Group':
                    _auxList = []
                    _auxList.append(key)
                    for sub_item in current_dict[key].keys():
                        _auxList.append(current_dict[key][sub_item])
                    ws.append(_auxList)
            ws.append([''])
            ws.append([''])
        
        wb.save(output)
        st.download_button(label= 'Download Report', data=output, file_name=f'{file_name}'+f'{_fileName}')
st.subheader('Generate Efficiency Report')
file_name = st.text_input('File Name', placeholder='Bolton')
uploadedfile = st.file_uploader('Upload Schedule File', type= 'xlsx')

if uploadedfile:    
    builder = TableBuilder(uploadedfile)

    builder.buildTableFile()
    #builder.tablesList

    #fs = FullSchedule(uploadedfile)
    #fs.insert_extra_columns()
    #prefGroupList = ['Corridor']
    #fs.build_printable_table(prefGroupList)

    

    #dicio = fs.finalTableDict

    #for item in dicio.items():
        #print(item)

    