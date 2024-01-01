import pandas as pd
#Fuzzywuzzy allows you to check strings for similarities 
from fuzzywuzzy import fuzz


#Get the student and alumni excel file
studentXls = pd.ExcelFile("students24.xlsx")
alumniXls = pd.ExcelFile("alumni24.xlsx")

#Finds the major of the given student
def findStudentMajor(student):
    #Students could be double major so use a list
    majors = []
    #For all the columns in the students row
    for c in list(student.index):
        #Check conditions and if it is a major then add to list
        if 'AA' <= c <= 'AR' and pd.notnull(student[c]) and student[c] != "Other:":
            majors.append(student[c])
    return majors

#Same as alumni but with different rows
def findAlumniMajor(alumni):
    majors = []
    for c in list(alumni.index):
        if 'Z' <= c <= 'AQ' and pd.notnull(alumni[c]) and alumni[c] != "Other:":
            majors.append(alumni[c])
    return majors

#Given a student career finds a matching alumni using FuzzyWuzzy
def findMatchingAlumniByCareer(studentCareer):
    maxRatio = -1
    matchedAlumni = None
    #Searches all alumni rows starting at row 9
    for _, alumni in alumniXls.iloc[8:].iterrows():
        alumniCareer = alumni['Q31_1_TEXT']
        ratio = fuzz.token_set_ratio(studentCareer.lower(), alumniCareer.lower())

        # Update the maximum similarity ratio(Percentage) and the matching alumni if a higher ratio is found
        if ratio > max_ratio:
            max_ratio = ratio
            matching_alumni = alumni if ratio >= 60  else None
    return matching_alumni

#Matches a student's pre-profesional track by the career outcome Ex. pre-med with doctor
def findByPreProfessionalTrack(studentTrack):
    #If pre-med and not paired by career then pair with any doctor
    if studentTrack == "Pre-Medicine":
        for _, alumni in alumniXls.iloc[8:].iterrows():
            alumniDegree = alumni['Q30']
            if alumniDegree == "Doctor of Medicine":
                return alumni
    #If pre-law then pair with attorney/lawyer
    if studentTrack == "Pre-Law":
        for _, alumni in alumniXls.iloc[8:].iterrows():
            alumniCareer = alumni['Q31_1_TEXT']
            #checks whether the alumni put lawyer/attorney
            if fuzz.token_set_ratio("lawyer", alumniCareer.lower()) >= 60 or fuzz.token_set_ratio("attorney", alumniCareer.lower()):
                return alumni

#Iterate over the students and match based on information
for _, student in studentXls.iloc[9:].iterrows():
    #student_name =
    student_major = findStudentMajor(student)
     