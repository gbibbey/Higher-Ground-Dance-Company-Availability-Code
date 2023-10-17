from xlrd import open_workbook
import pandas as pd

fullSheetWorkbook = open_workbook("FullResponseSheet.xls")   
fullSheet = fullSheetWorkbook.sheet_by_index(0)
requested = dict({})

# used for cross-referencing
for i in range(1, fullSheet.nrows):
    name = fullSheet.cell_value(i, 2)
    requested[name] = []
    for j in range(12, 57):  
        if fullSheet.cell_value(i, j) != "":
            requested[name].append(fullSheet.cell_value(0, j))

danceTimeSheet = open_workbook("dancesAndTimes.xls") 
danceTimes = danceTimeSheet.sheet_by_index(0)

# Reads in dancer table
# 0 as start and end for any day means no availability

availabilityByDance = dict({})
dancerDict = dict({})
availabilityS = 0
availabilityE = 0

for i in range(1, danceTimes.nrows):
    songName = danceTimes.cell_value(i, 0)
    day = danceTimes.cell_value(i, 1)
    time = danceTimes.cell_value(i, 2)

    # add the song to the dict
    if songName not in availabilityByDance:
        availabilityByDance[songName] = []

    for dancerIndex in range(1, fullSheet.nrows):
        dancerName = fullSheet.cell_value(dancerIndex, 2)

        if dancerName not in dancerDict:
            dancerDict[dancerName] = []

        index = 1
        if day == "Sunday":
            index = 57
        elif day == "Monday":
            index = 58
        elif day == "Tuesday":
            index = 59
        elif day == "Wednesday":
            index = 60
        elif day == "Thursday":
            index = 61

        # Split stringified array of dancer UNavailability
        dancerUnavailability = fullSheet.cell_value(dancerIndex, index)
        print(dancerName)
        print(dancerUnavailability)
        
        if time not in dancerUnavailability and songName not in dancerDict[dancerName]:
            availabilityByDance[songName].append(dancerName)
            if songName in requested[dancerName]:
                dancerDict[dancerName].append(songName + "*")
            else:
                dancerDict[dancerName].append(songName)
        else:
            dancerDict[dancerName].append(" ")

            
#print(availabilityByDance)
#print(dancerDict)

def highlight_requested(val):
    color = 'green' if "*" in val else 'white'
    return 'background-color: {}'.format(color)


df = pd.DataFrame(dancerDict) 
df.style.applymap(highlight_requested).to_excel('highlightedAvailability.xlsx', engine='openpyxl')
