# Import Libraries
import pandas as pd
import csv

# Read files

ci = pd.read_excel('Consultant Requests.xlsx', header=0) #ci is Consultant Information
wc = pd.read_excel('Director Requests.xlsx', header=0) #wc is Writing Center

# Classes
class Consultant:  #Class that holds consultant info
    def __init__(self, Name, Year, Field_Of_Study, Hours_Wanted, Times_Available):
      self.Name = Name                           ##Single Value
      self.Year = Year                           ##Single Value
      self.Hours_Wanted = int(Hours_Wanted)      ##Single Value
      self.Field_Of_Study = Field_Of_Study       ##List
      self.Times_Available = Times_Available     ##List
      self.NumberOfShifts = 0 ##List
BlankWorker = Consultant("-1", "", [],0, [])

class Shift:  #Class that holds Shift info
    def __init__(self, hour, priority):
      self.hour = hour                         ##Single Value
      self.priority = priority                 ##Single Value
      self.workerNames = []                    ##List

# Non-Main Functions
def CreateConsultantList():
    ConsultantList = []
    for index, row in ci.iterrows(): # creating classes containing individual consultants
        # combine first name and last name into a single data value
        firstName = row['First Name']
        lastName = row['Last Name']
        fullName = firstName.strip() + " " + lastName.strip() # full name

        # create availability list
        sunday = row['Sunday']
        monday = row['Monday']
        tuesday = row['Tuesday']
        wednesday = row['Wednesday']
        thursday = row['Thursday']
        friday = row['Friday']

        week = [sunday, monday, tuesday, wednesday, thursday, friday]
        weekNames = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
        TempTimesAvailableList = []
        for i in range(len(week)):
            if not isinstance(week[i], float):
                week[i] = week[i].strip()
                PreTempTimesAvailableList=[weekNames[i]+time.strip() for time in week[i].split(",")]
                for index in range(len(PreTempTimesAvailableList)):
                    if "->" in PreTempTimesAvailableList[index]:
                        TempTimesAvailableList.append(PreTempTimesAvailableList[index])
                        # print(PreTempTimesAvailableList[index])

        # create major(s)/minor(s) list
        Field_of_Study = row['Majors/Minors']
        Field_of_Study.strip()
        TempFieldStudyList=[i.strip() for i in Field_of_Study.split(",")]

        # create requested hours
        requested_hours = row['Hours']
        if row['Hours'] == '10+': # checks if requested hours are 10+ and assigns a integer number of hours
            requested_hours = 10

        # create consultant
        worker = Consultant(Name=fullName, Year=row['Year'], Hours_Wanted=requested_hours, Times_Available=TempTimesAvailableList, Field_Of_Study=TempFieldStudyList) 
        ConsultantList.append(worker)
    return ConsultantList

def CreateShiftInfo():
    # Shift consultant information
    for index, row in wc.iterrows():
        if (row['Multiple Shifts'].strip() == 'Yes'):
            Multiple_Shifts = True
        else:
            Multiple_Shifts = False
        if (row['Mix Majors'].strip() == 'Yes'):
            Mix_Majors = True
        else:
            Mix_Majors = False
        if (row['Mix Years'].strip() == 'Yes'):
            Mix_Years = True
        else:
            Mix_Years = False

        Shift_Min = row['Shift Minimum']
        if Shift_Min == '10+': # checks if minimum is 10+ and assigns a integer number of hours
            Shift_Min = 10
        Shift_Max = row['Shift Maximum']
        if Shift_Max == '10+': # checks if maximum is 10+ and assigns a integer number of hours
            Shift_Max = 10

        Choices_List = [Multiple_Shifts, Mix_Majors, Mix_Years, Shift_Min, Shift_Max]


        # Shift Information
        Shift_List = []
        weekNames = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
            # open hours

        o_sunday = row['Sunday Hours']
        o_monday = row['Monday Hours']
        o_tuesday = row['Tuesday Hours']
        o_wednesday = row['Wednesday Hours']
        o_thursday = row['Thursday Hours']
        o_friday = row['Friday Hours']
            # busy hours

        b_sunday = row['Sunday Busy Hours']
        b_monday = row['Monday Busy Hours']
        b_tuesday = row['Tuesday Busy Hours']
        b_wednesday = row['Wednesday Busy Hours']
        b_thursday = row['Thursday Busy Hours']
        b_friday = row['Friday Busy Hours']
            # quiet hours

        q_sunday = row['Sunday Quiet Hours']
        q_monday = row['Monday Quiet Hours']
        q_tuesday = row['Tuesday Quiet Hours']
        q_wednesday = row['Wednesday Quiet Hours']
        q_thursday = row['Thursday Quiet Hours']
        q_friday = row['Friday Quiet Hours']

        open_shifts = [o_sunday, o_monday, o_tuesday, o_wednesday, o_thursday, o_friday]
        busy_shifts = [b_sunday, b_monday, b_tuesday, b_wednesday, b_thursday, b_friday]
        quiet_shifts = [q_sunday, q_monday, q_tuesday, q_wednesday, q_thursday, q_friday]

        for i in range(len(open_shifts)):
            open_shifts[i] = open_shifts[i].strip().split(",")
            busy_shifts[i] = busy_shifts[i].strip().split(",")
            quiet_shifts[i] = quiet_shifts[i].strip().split(",")


            for time in range(len(open_shifts[i])):
                #print(weekNames[i]+open_shifts[i][time])

                if open_shifts[i][time].strip() != "":
                    if open_shifts[i][time] in busy_shifts[i]:
                        work_shift = Shift(weekNames[i]+open_shifts[i][time].strip(), 1)
                    elif open_shifts[i][time] in quiet_shifts[i]:
                        work_shift = Shift(weekNames[i]+open_shifts[i][time].strip(), 3)
                    else:
                        work_shift = Shift(weekNames[i]+open_shifts[i][time].strip(), 2)

                    Shift_List.append(work_shift)
                    #print(work_shift.hour, " ", work_shift.priority)

    return Shift_List, Choices_List


def ScheduleTrimmer(Shift_List,Shift,Max_Workers,CountUpper,Multiple_Shifts,Mix_Majors,Mix_Years):
  NumExtraWorkers = len(Shift_List[Shift].workerNames)-Max_Workers #determines if there's extra workers in a shift
  if (NumExtraWorkers>0):
      LargestNumberOfShifts=0
      NumberOfUpperClassmen=0
      TempKickList=[]
      for worker in Shift_List[Shift].workerNames:
          if worker.NumberOfShifts > worker.Hours_Wanted:
              TempKickList.append(worker) #apending all workers scheduled for to many shifts
              LargestNumberOfShifts = max(LargestNumberOfShifts , worker.NumberOfShifts)
          if(worker.Year == "Senior" or worker.Year == "Junior"):
              NumberOfUpperClassmen+=1

      #Years
      if Mix_Years:
         if(CountUpper and NumberOfUpperClassmen == 1): #Removes upperclass if only one
           for staff in TempKickList:
              if (staff.Year == "Senior" or staff.Year == "Junior"):
                  TempKickList.remove(staff)

      #Consecutive Shifts
      if Multiple_Shifts:
          for staff in TempKickList:
              if NumExtraWorkers>0:
                  if staff in Shift_List[Shift-1].workerNames or staff in Shift_List[min(Shift+1,len(Shift_List)-1)].workerNames:
                      if len(TempKickList) > NumExtraWorkers:
                        #Removes Staff from temp list
                        TempKickList.remove(staff)
      else:
          for staff in TempKickList:
            if NumExtraWorkers>0:
                 if staff in Shift_List[Shift-1].workerNames or staff in Shift_List[min(Shift+1,len(Shift_List)-1)].workerNames:
                    #Removes Staff from shift
                    NumExtraWorkers -=1
                    staff.NumberOfShifts-=1
                    TempKickList.remove(staff)
                    Shift_List[Shift].workerNames.remove(staff)

      #Mixing Majors
      matches_found = True
      if Mix_Majors:
         while (len(TempKickList)>0 and NumExtraWorkers>0 and matches_found):
            MostMatchesOverall = -1
            CurrentLargestFieldWorker = BlankWorker
                        
            for staff in TempKickList:
                MostMatches = -1
                for FieldStudy in staff.Field_Of_Study:##Possible to have more than 1 major
                     MajorMatches = 0
                     for otherstaff in Shift_List[Shift].workerNames:
                        if FieldStudy in otherstaff.Field_Of_Study:
                            MajorMatches+=1
                     if MajorMatches > MostMatches:
                      MostMatches = MajorMatches    
                if MostMatches >= MostMatchesOverall:
                  MostMatchesOverall = MostMatches
                  CurrentLargestFieldWorker = staff

            if MostMatchesOverall > 1:#probably 1 and not 0
                #Removes Staff from shift
                NumExtraWorkers -=1
                CurrentLargestFieldWorker.NumberOfShifts-=1
                TempKickList.remove(CurrentLargestFieldWorker)
                Shift_List[Shift].workerNames.remove(CurrentLargestFieldWorker)
            else:
                matches_found = False

      #Largest Differences
      while (len(TempKickList)>0 and NumExtraWorkers>0):
        CurrentLargestFieldWorker = BlankWorker
        biggest_diff = 0
        for staff in TempKickList:
            if (staff.NumberOfShifts - staff.Hours_Wanted) > biggest_diff:
                CurrentLargestFieldWorker = staff
                biggest_diff = staff.NumberOfShifts - staff.Hours_Wanted

        #Removes Staff from shift
        NumExtraWorkers -=1
        CurrentLargestFieldWorker.NumberOfShifts-=1
        TempKickList.remove(CurrentLargestFieldWorker)
        Shift_List[Shift].workerNames.remove(CurrentLargestFieldWorker)


def CreateOutputFile (Shift_List,ConsultantList):
  OutFileName="Schedule.csv"
  with open (OutFileName, "w", encoding='utf-8') as OutFile:
    Columns = ["Time", "Workers", "Number Of Workers"]
    writer = csv.DictWriter(OutFile, fieldnames=Columns)
    writer.writeheader()
    for Shift in Shift_List:
        TempNameList=[]
        for name in Shift.workerNames:
            TempNameList.append(name.Name)
        writer.writerow({"Time": Shift.hour, "Workers": TempNameList, "Number Of Workers": len(TempNameList)})
    OutFile.close()

# Main Function
def Main():

    # Program set up
    ConsultantList = CreateConsultantList()
    #for i in ConsultantList: # prints list of consultants and attached info
        #print(i.Name, i.Year, i.Hours_Wanted)
        #print(i.Field_Of_Study, i.Times_Available)
    Shift_List, Choices_List = CreateShiftInfo()
    #for i in Shift_List: # prints list of available shifts
        #print(i.hour, i.priority)
    #for i in Choices_List: # prints director's choices
       # print(i)

    # Scheduler Program
    Max_Workers = Choices_List[4]
    Min_Workers = Choices_List[3]
    for worker in ConsultantList: #Workers get distributed to available work times
        for TimeAvailable in worker.Times_Available: # runs through the availability of worker
            for work_shift in Shift_List: # runs through all avaialable shifts
                if (len(work_shift.workerNames)< Max_Workers): # checks that shift is not already full
                    if worker.NumberOfShifts < (worker.Hours_Wanted*2): # limits worker to only twice the number of shifts requested
                        if TimeAvailable == work_shift.hour: # checks if worker's availability is the same as an open shift
                            work_shift.workerNames.append(worker) # adds worker to shift
                            worker.NumberOfShifts+=1 # adds to worker's total number of shifts

    for Shift in (range(len(Shift_List))): #Tries to reduce shifts to under Shift_Maximum
        ScheduleTrimmer(Shift_List,Shift,Choices_List[4],True,Choices_List[0],Choices_List[1],Choices_List[2])

    for Shift in range(len(Shift_List)): 
        priority = 3 # priority scales from 1 to 3, with 3 being the least busy shifts and 1 the busiest
        while (priority>1): # tries to reduce amount of workers in least busy and average shifts to Shift_Minimum
            for Shift in (range(len(Shift_List))):
                if Shift_List[Shift].priority == priority:
                    ScheduleTrimmer(Shift_List,Shift,Choices_List[3],True,Choices_List[0],Choices_List[1],Choices_List[2])
            priority-=1
        while (priority==1): # tries to reduce amount of workers in busiest shifts to under Shift_Maximum
            for Shift in range(len(Shift_List)): #Tries to reduce shifts to under Shift_Maximum
                ScheduleTrimmer(Shift_List,Shift,Choices_List[4],True,Choices_List[0],Choices_List[1],Choices_List[2])
            priority-=1

    for worker in ConsultantList:
        if worker.Hours_Wanted > worker.NumberOfShifts: # checks if worker has less shifts than requested
            print(worker.Name, " has less than their requested hours!")
            for TimeAvailable in worker.Times_Available: # runs through the availability of worker
                if worker.Hours_Wanted > worker.NumberOfShifts: # checks if worker still needs more shifts
                    for work_shift in Shift_List: # runs through all avaialable shifts
                        if (len(work_shift.workerNames)< Min_Workers): # looks for shifts that do not have the minimum workers
                            if TimeAvailable == work_shift.hour: # checks if worker's availability is the same as an open shift
                                work_shift.workerNames.append(worker) # adds worker to shift
                                worker.NumberOfShifts+=1 # adds to worker's total number of shifts

            if worker.Hours_Wanted > worker.NumberOfShifts: # checks if worker still needs more shifts
                for TimeAvailable in worker.Times_Available: # runs through the availability of worker
                    for work_shift in Shift_List: # runs through all avaialable shifts
                        if (len(work_shift.workerNames)< Max_Workers): # checks that shift is not already full
                            if (work_shift.priority == 3): # checks if shift is a busy shift
                                if TimeAvailable == work_shift.hour: # checks if worker's availability is the same as an open shift
                                    work_shift.workerNames.append(worker) # adds worker to shift
                                    worker.NumberOfShifts+=1 # adds to worker's total number of shifts
                            if (work_shift.priority == 2) and (len(work_shift.workerNames)<(Max_Workers-1)): # checks if shift is an average shift and that placing the worker in the shift doesn't make the shift full
                                if TimeAvailable == work_shift.hour: # checks if worker's availability is the same as an open shift
                                    work_shift.workerNames.append(worker) # adds worker to shift
                                    worker.NumberOfShifts+=1 # adds to worker's total number of shifts

        if worker.Hours_Wanted < worker.NumberOfShifts: # checks if worker has more shifts than requested
            print(worker.Name, " has more than their requested hours!")
            for work_shift in Shift_List: # tries to remove worker's extra shifts
                if worker.Hours_Wanted < worker.NumberOfShifts: # checks worker still needs to be removed from shifts
                    if worker in work_shift.workerNames: # checks if worker is in shift
                        if (len(work_shift.workerNames)> Min_Workers): # checks that if removing the worker would put the shift below minimum shift requirements
                            work_shift.workerNames.remove(worker) # removes worker from shift
                            worker.NumberOfShifts-=1 # reduces worker's number of shifts
        if worker.Hours_Wanted == worker.NumberOfShifts:
            print(worker.Name, " has their requested hours!")


    CreateOutputFile (Shift_List,ConsultantList) #Creates the "Schedule.csv" file



Main()
    
