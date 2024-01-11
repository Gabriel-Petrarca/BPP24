import pandas as pd
#Fuzzywuzzy allows you to check strings for similarities 
from fuzzywuzzy import fuzz


#Get the student and alumni excel file
studentXls = pd.read_excel("BPP24Student.xlsx")
alumniXls = pd.read_excel("BPP24Alumni.xlsx")

#Alumni helper functions
def getAlumniName(alumni):
    if isinstance(alumni['Q3'], str):
        return alumni['Q3'] + ' ' + alumni['Q4']
def getAlumniEmail(alumni):
    return alumni['Q8']
def getAlumniCareer(alumni):
    return alumni['Q31_1_TEXT']
def getAlumniPronouns(alumni):
    return alumni['Q6']
def getAlumniBlurb(alumni):
    return alumni['Q32']
#Finds the major of the given student
def findStudentMajor(student):
    #Students could be double major so use a list
    majors = []
    #For all the columns in the students row
    for c in studentXls.columns[studentXls.columns.get_loc('Q13'):studentXls.columns.get_loc('Q29_10_TEXT') + 1]:
        #Check conditions and if it is a major then add to list
        if pd.notnull(student[c]) and student[c] != "Other:":
            majors.append(student[c])
    return majors

#Same as alumni but with different cols
def findAlumniMajor(alumni):
    majors = []
    for c in studentXls.columns[studentXls.columns.get_loc('Q13'):studentXls.columns.get_loc('Q29_10_TEXT') + 1]:
        if pd.notnull(alumni[c]) and alumni[c] != "Other:":
            majors.append(alumni[c])
    return majors

#Given a student career finds a matching alumni using FuzzyWuzzy
def findMatchingAlumniByCareer(studentCareer):
    maxRatio = -1
    matchingAlumni = None
    #Searches all alumni rows starting at row 9
    for _, alumni in alumniXls.iloc[8:].iterrows():
        alumniCareer = getAlumniCareer(alumni)
        if isinstance(studentCareer, str) and isinstance(alumniCareer, str):
            ratio = fuzz.token_set_ratio(studentCareer.lower(), alumniCareer.lower())
        else:
            continue
        # Update the maximum similarity ratio(Percentage) and the matching alumni if a higher ratio is found
        if ratio > maxRatio:
            maxRatio = ratio
            matchingAlumni = alumni if ratio >= 60  else None
    return matchingAlumni

#Matches student based on majors
def findByMajor(studentMajors):
    alumniMatch = None
    for _, alumni in alumniXls.iloc[8:].iterrows():
        alumniMajors = findAlumniMajor(alumni)
        if studentMajors == alumniMajors:
            alumniMatch = alumni
        else: 
            for major in studentMajors:
                if major in alumniMajors:
                    alumniMatch = alumni
    return alumniMatch
        
#Matches a student's pre-profesional track by the career outcome Ex. pre-med with doctor
def findByPreProfessionalTrack(studentTrack):
    #If pre-med and not paired by career then pair with any doctor
    alumniMatch = None
    if studentTrack == "Pre-Medicine":
        for _, alumni in alumniXls.iloc[8:].iterrows():
            alumniDegree = alumni['Q30']
            if alumniDegree == "Doctor of Medicine":
                alumniMatch = alumni
                return alumniMatch
    #If pre-law then pair with attorney/lawyer
    if studentTrack == "Pre-Law":
        for _, alumni in alumniXls.iloc[8:].iterrows():
            alumniCareer = getAlumniCareer(alumni)
            #checks whether the alumni put lawyer/attorney

            if isinstance(studentCareer, str) and isinstance(alumniCareer, str):
                if fuzz.token_set_ratio("lawyer", alumniCareer.lower()) >= 60 or fuzz.token_set_ratio("attorney", alumniCareer.lower()) >= 60:
                    alumniMatch = alumni
                    return alumniMatch
    return alumniMatch
#Iterate over the students and match based on information
matches = []
for _, student in studentXls.iloc[9:].iterrows():
    #Concatenates first and last names
    if isinstance(student['Q3'], str) and isinstance(student['Q4'], str):
        studentName =  student['Q3'] + ' ' + student['Q4']
    studentMajors = findStudentMajor(student)
    studentCareer = student['Q31_1_TEXT']
    studentEmail = student['Q8']
    studentTrack = student['Q30']
    studentPronouns = student['Q6']

    if studentMajors == []:
        studentXls = studentXls[studentXls['Q3'] + ' ' + studentXls['Q4'] != studentName]
    #Finds if there is a match by career
    alumniMatch = findMatchingAlumniByCareer(studentCareer)
    
    #If there is a match update the new excel sheet and remove student and alumni from the old one
    if alumniMatch is not None:
        alumniName = getAlumniName(alumniMatch)
        alumniEmail = getAlumniEmail(alumniMatch)
        alumniCareer = getAlumniCareer(alumniMatch)
        alumniPronouns = getAlumniPronouns(alumniMatch)
        alumniBlurb = getAlumniBlurb(alumniMatch)
        matches.append([studentName, studentEmail, studentPronouns, studentCareer, alumniCareer, alumniName, alumniEmail, alumniPronouns, alumniBlurb])

        #Removes the row containing student and alumni email / keeps all rows except student and alumni email
        studentXls = studentXls[studentXls['Q3'] + ' ' + studentXls['Q4'] != studentName]
        alumniXls = alumniXls[alumniXls['Q3'] + ' ' + alumniXls['Q4'] != alumniName]
    #If not matched by career search by pre-track
    if alumniMatch is None:
        if studentTrack == "Pre-Medicine" or studentTrack == "Pre-Law":
            alumniMatch = findByPreProfessionalTrack(studentTrack)
        if alumniMatch is not None:
            alumniName = getAlumniName(alumniMatch)
            alumniEmail = getAlumniEmail(alumniMatch)
            alumniCareer = getAlumniCareer(alumniMatch)
            alumniPronouns = getAlumniPronouns(alumniMatch)
            alumniBlurb = getAlumniBlurb(alumniMatch)
            matches.append([studentName, studentEmail, studentPronouns, studentCareer, alumniCareer, alumniName, alumniEmail, alumniPronouns, alumniBlurb])

            studentXls = studentXls[studentXls['Q3'] + ' ' + studentXls['Q4'] != studentName]
            alumniXls = alumniXls[alumniXls['Q3'] + ' ' + alumniXls['Q4'] != alumniName]
    #Else search by major
    if alumniMatch is None:
        alumniMatch = findByMajor(studentMajors)
        if alumniMatch is not None:
            alumniName = getAlumniName(alumniMatch)
            alumniEmail = getAlumniEmail(alumniMatch)
            alumniCareer = getAlumniCareer(alumniMatch)
            alumniMajors = findAlumniMajor(alumniMatch)
            alumniPronouns = getAlumniPronouns(alumniMatch)
            alumniBlurb = getAlumniBlurb(alumniMatch)

            if alumniMajors != []:
                matches.append([studentName, studentEmail, studentPronouns, ', '.join(studentMajors), ', '.join(alumniMajors), alumniName, alumniEmail, alumniPronouns, alumniBlurb])
                alumniXls = alumniXls[alumniXls['Q3'] + ' ' + alumniXls['Q4'] != alumniName]
                studentXls = studentXls[studentXls['Q3'] + ' ' + studentXls['Q4'] != studentName]
            else:
                alumniXls = alumniXls[alumniXls['Q3'] + ' ' + alumniXls['Q4'] != alumniName]

            
            
matchedPairsDF = pd.DataFrame(matches, columns=['Student Name', 'Student Email', 'Student Pronouns', 'Student Matched By', 'Alumni Matched By', 'Alumni Name', 'Alumni Email', 'Alumni Pronouns', 'Alumni Bio'])

with pd.ExcelWriter('BPP2324.xlsx') as writer:
    matchedPairsDF.to_excel(writer, sheet_name = 'Pairs', index = False)
    studentXls.to_excel(writer, sheet_name = 'students left', index = False)
    alumniXls.to_excel(writer, sheet_name = 'alumni left', index = False)
