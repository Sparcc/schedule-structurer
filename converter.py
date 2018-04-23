import re
from openpyxl import load_workbook
wb = load_workbook("schedule.xlsx")
ws = wb.active
startingCells = 4
maxCells = 261
regexCode = '(\d{1,3}) (\d{0,3}%) (.+) (\d{0,3} days{0,1}|\d{0,1}.\d{0,1}.\d{0,1} days{0,1}) (\S{3} \d{1,2}\/\d{1,2}\/\d{1,2}) (\S{3} \d{1,2}\/\d{1,2}\/\d{1,2})( \d{1,3}|)(.*){0,1}'
class ScheduleEntry:
    ID =""
    Completion = ""
    Task = ""
    Duration = ""
    Start = ""
    Finish = ""
    Predecessors = ""
    Resources = "Not Defined"

schedule = []
i=startingCells
while i < maxCells:
    text = ws["A"+str(i)].value
    #print(text)
    result = re.search(regexCode, text)
    if result:
        #print("Match Found")
        resultGroups = result.groups()
        #print(resultGroups)
        if len(resultGroups) > 2:
            scheduleEntry = ScheduleEntry()
            scheduleEntry.ID = str(resultGroups[0])
            scheduleEntry.Completion = str(resultGroups[1])
            scheduleEntry.Task = str(resultGroups[2])
            scheduleEntry.Duration = str(resultGroups[3])
            scheduleEntry.Start = str(resultGroups[4])
            scheduleEntry.Finish = str(resultGroups[5])
            
            if len(resultGroups) > 5:
                scheduleEntry.Predecessors = resultGroups[6]
                resultGroups[6]
                if len(resultGroups) > 6:
                    scheduleEntry.Resources = resultGroups[7]
                    resultGroups[7]
            schedule.append(scheduleEntry)
            
    else:
        print("No Match")
    i+=1
    

for value in schedule:
    searchForMe = '(.*Thomas.*)'
    result = re.search(searchForMe,value.Resources)
    if result:
        print(#value.ID+", "+
        #value.Completion+", "+
        value.Task+", "+
        value.Duration+", "+
        value.Start+", "+
        value.Finish+", "+
        #value.Predecessors+", "+
        value.Resources)