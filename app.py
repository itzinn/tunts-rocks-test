import gspread
from oauth2client.service_account import ServiceAccountCredentials
import math

scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']

creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
client = gspread.authorize(creds)

spreadsheet = client.open_by_url('https://docs.google.com/spreadsheets/d/1UhNv6OinkjJSL0nWw66E7Cf-IcJMDqrbZXsy05LALwo/edit#gid=0')

sheet = spreadsheet.get_worksheet(0)

#skip the first three rows of the spreadsheet
header_offset = 3

#retrieve the values from the columns to be processed
aluno_values = sheet.col_values(2)[header_offset:]
faltas_values = sheet.col_values(3)[header_offset:]
p1_values = sheet.col_values(4)[header_offset:]
p2_values = sheet.col_values(5)[header_offset:]
p3_values = sheet.col_values(6)[header_offset:]

print(aluno_values)

'''
receives parameters faltas, p1, p2, p3
return the student's situation and their average grade.
'''
def getSituation(faltas, p1, p2, p3):
    situation = ""
    m = 0

    if(faltas/60 >= 0.25):
        situation = "Reprovado por Falta"
    else:
        m = (p1 + p2 + p3)/3

        if(m < 50):
            situation = "Reprovado por Nota"
        elif(50 <= m < 70):
            situation = "Exame Final"
        elif(m >= 70):
            situation = "Aprovado"

    return situation, m

'''
receives the student's situation and their average grade
returns the minimum grade the student needs to achieve on the final exam to pass.
'''
def getNAF(situation, m):
    naf = 0

    if(situation == "Exame Final"):
            naf = math.ceil(100 - m)

    return naf

#loop that iterates through the spreadsheet and analyzes the situation of each student
for i in range(len(aluno_values)):

    situation, m = getSituation(int(faltas_values[i]), int(p1_values[i]), int(p2_values[i]), int(p3_values[i]))
    
    naf = getNAF(situation, m)
    
    print(f"Aluno: {aluno_values[i]}\tSituacao: {situation}\tNAF: {naf}")

    #write the student's situation and NAF in columns 7 and 8 of the spreadsheet
    sheet.update_cell(i+1+header_offset, 7, situation)
    sheet.update_cell(i+1+header_offset, 8, naf)

