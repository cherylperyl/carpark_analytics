import pandas as pd
import numpy as np
import datetime
import warnings
warnings.filterwarnings('ignore')

# FIRST OPERATION

# variables
carparkNo =  input("Enter the carpark number:")
month =  "JUNE"
year = "2022"
reportName = "HISTORY TRANSACTION"
fileFormat = "csv"
path = "input_files" + "/CARPARK" + carparkNo + " " + month + " " + year  + " " + reportName + "." + fileFormat

# read raw file
df = pd.read_csv(path)
df["TIME"] = pd.to_datetime(df["TIME"], dayfirst=True)
df.sort_values("TIME", inplace=True)

# remove the random spaces in STATION attribute
def strip(station):
    return station.strip()
df["STATION"] = df["STATION"].apply(strip)

# collect relevant data needed to match records in a list (using a list so i can iterate through easily)
data = []
for index, row in df.iterrows():
    id, time, station, iu, cashcard, amount = row["S/NO."], row["TIME"], row["STATION"], row["IU"], row["CASHCARD"], row["AMOUNT"]
    data.append({"id":id,
                "time":time, 
                "station":station,
                "iu":iu,
                "amount": amount,
                "cashcard": cashcard})

# match exits to entries
for i in range(len(data)):
    record = data[i]
    if record["station"] == "Entry":
        for j in range(i+1, len(data)):
            exit_record = data[j]
            if exit_record["station"] == "Exit" and exit_record["iu"] == record["iu"] and exit_record["time"] >= record["time"]:
                data[i]["exitTime"] = exit_record["time"]
                data[i]["amount"] =  exit_record["amount"]
                data[i]["cashcard"] = exit_record["cashcard"]
                break

# create a new dict to store the exit times by S/NO.
exitTimes = {}
for record in data:
    if record["station"] == "Entry":
        if "exitTime" not in record:
            record["exitTime"] = np.NaN
        exitTimes[record["id"]] = {"exitTime": record["exitTime"],
                                    "amount": record["amount"],
                                    "cashcard": record["cashcard"]}

# filter for Entry records
entry_df = df[df["STATION"] == "Entry"]

# use a list to order the exit times according to the sequence of records in entry_df (using S/NO. to identify the record)
exitTimesList = []
exitAmountList = []
exitCashcardList = []
for i in entry_df["S/NO."]:
    exitTimesList.append(exitTimes[i]["exitTime"])
    exitAmountList.append(exitTimes[i]["amount"])
    exitCashcardList.append(exitTimes[i]["cashcard"])

# add exit times to the dataset
entry_df["EXITDATETIME"] = exitTimesList
entry_df["AMOUNT"] = exitAmountList
entry_df["CASHCARD"] = exitCashcardList

# remove unnecessary column
entry_df.drop(columns=["STATION"], inplace=True)

# rename TIME column to ENTRYTIME
entry_df.rename(columns={"TIME":"ENTRYDATETIME"}, inplace=True)

# save to CSV
path =  "output_files" + "/entryExitMatched/CARPARK" + carparkNo + "_entryExitMatched.csv"
entry_df.to_csv(path, index=False)
print("File generated successfully and can be found in: "+path)

# SECOND OPERATION

# read file
entry_df = pd.read_csv(path)
entry_df["ENTRYDATETIME"] = pd.to_datetime(entry_df["ENTRYDATETIME"], dayfirst=True)
entry_df["EXITDATETIME"] = pd.to_datetime(entry_df["EXITDATETIME"], dayfirst=True)

# generate columns for each hour
option = "Hours"
multiple = "1"
frequency = {"Seconds": "S",
            "Minutes": "min",
            "Hours": "H"}
columns = pd.date_range(start="00:00:00", end="23:59:59", freq=multiple+frequency[option])

# round down datetime
def roundTime(start, end, option):
    if option == "Hours":
        return start.replace(microsecond=0, second=0, minute=0), end.replace(microsecond=0, second=0, minute=0)
    elif option == "Minutes":
        return start.replace(microsecond=0, second=0), end.replace(microsecond=0, second=0)
    elif option =="Seconds":
        return start.replace(microsecond=0), end.replace(microsecond=0)
    else:
        print("Invalid Option!")

# add time ranges for each record
timeRanges = []
for index, row in entry_df.iterrows():
    try:
        startTime, endTime = roundTime(row["ENTRYDATETIME"], row["EXITDATETIME"], option)
        timeRanges.append(pd.date_range(startTime, endTime, freq=multiple+frequency[option]))
    except ValueError:
        timeRanges.append(np.NaN)
entry_df["timeRange"] =  timeRanges

# if hour is found in the record's time range, 1 is added to its hour column
for time in columns:
    specific_time = []
    for index, row in entry_df.iterrows():
        time_found = False
        timeRanges = row["timeRange"]
        if type(timeRanges) != float:
            for timestamp in timeRanges:
                if time.time() == timestamp.time():
                    specific_time.append(1)
                    time_found = True
                    break
            if time_found == False:
                specific_time.append(0)
        else:
            specific_time.append(0)
    entry_df[time.time()] = specific_time

# extract date from datetime column
def getDate(dtObject):
    if type(dtObject) != float:
        return str(dtObject).split()[0]

entry_df["DATE"] = entry_df["ENTRYDATETIME"].apply(getDate)
byDay_df = entry_df.groupby("DATE").sum()
byDay_df.drop(columns=["S/NO.", "VEHICLE"], inplace=True)
byDay_df.reset_index(inplace=True)

# add days of the week column
def getWeekday(date):
    dateComponents = date.split('-')
    dtObject = datetime.date(year=int(dateComponents[0]), month=int(dateComponents[1]), day=int(dateComponents[2]))
    return dtObject.weekday()

byDay_df["DAY"] = byDay_df["DATE"].apply(getWeekday)

byWeekday_df = byDay_df.groupby("DAY").mean().reset_index().sort_values("DAY")

# convert number to days of the week
def convertToDaysOfTheWeek(num):
    weekdays = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    return weekdays[int(num)]

byWeekday_df["DAY"] = byWeekday_df["DAY"].apply(convertToDaysOfTheWeek)
byDay_df["DAY"] = byDay_df["DAY"].apply(convertToDaysOfTheWeek)

# create excel
excelName = "CARPARK" + carparkNo + "_heatmap" 

path = "output_files/heatmaps/" + excelName + ".xlsx"
writer = pd.ExcelWriter(path, engine='xlsxwriter')

# add data to excel sheet
byDay_df.to_excel(writer, sheet_name="byDay", index=False)

# Get sheet for conditional formatting 
worksheet = writer.sheets['byDay']

# Add conditional formatting
rowCount = str(byDay_df.shape[0]+1)
worksheet.conditional_format('B2:Y'+rowCount, {'type': '3_color_scale',
                                                'min_type': 'min',
                                                'mid_type': 'percent',
                                                'mid_value': 25,
                                                'max_type': 'max',
                                                'min_color': '#63BE7B',
                                                'mid_color': '#FFEB84',
                                                'max_color': '#F8696B',
                                                })

# add data to excel sheet
byWeekday_df.to_excel(writer, sheet_name="byWeekday", index=False)

# Get sheet for conditional formatting 
worksheet = writer.sheets['byWeekday']

# Add conditional formatting
rowCount = str(byWeekday_df.shape[0]+1)
worksheet.conditional_format('B2:Y'+rowCount, {'type': '3_color_scale',
                                                'min_type': 'min',
                                                'mid_type': 'percent',
                                                'mid_value': 25,
                                                'max_type': 'max',
                                                'min_color': '#63BE7B',
                                                'mid_color': '#FFEB84',
                                                'max_color': '#F8696B',
                                                })

writer.save()
print("Heatmap generated successfully and can be found in: " + path)
