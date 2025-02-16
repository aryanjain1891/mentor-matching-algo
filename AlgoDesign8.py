#This Algorithm Pairs up mentors on the basis of Startup's Preferences
import xlrd
xlrd.xlsx.ensure_elementtree_imported(False, None)
xlrd.xlsx.Element_has_iter = True
from collections import defaultdict
loc = ("//Users//abirabh//Documents//Top30MentorReal01.xlsx")
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)
rows = sheet.nrows
cols = sheet.ncols
Mend = {}
Startd = {}
for i in range(1,rows):
    l = [sheet.cell_value(i, 4)]
    x = sheet.cell_value(i, 0)
    x = x[1:]
    Mend[x] = l
print(Mend)
#Mend Stores the number of slots available for each mentor
loc2 = ("//Users//abirabh//Documents//Top30StartupPref.xlsx")
wb = xlrd.open_workbook(loc2)
sheet2 = wb.sheet_by_index(0)
rows2 = sheet2.nrows
cols2 = sheet2.ncols
l1 = []
l10 = []
for i in range(1,cols2):
    l00 = [int(sheet2.cell_value(1,i)),int(sheet2.cell_value(0, i))]
    l10.append(l00)
l10.sort(reverse = True)
#l10 contains startup preferences according to a preference order which Team Conquest decided on the basis
#of level of involvement of the startup
    
for i1 in l10:
    i = i1[1]-8
    l = []
    for j in range(1,rows2):
        x2 = sheet2.cell_value(j, i)
        if x2 != '':
            l.append(int(sheet2.cell_value(j, i)))
        else:
            break
    l.append(int(sheet2.cell_value(0, i)))
    l.append(0)
    l1.append(l)
#l1 has possible pairings for the mentors and startups
#finalpairup contains Final Matchings for Startups and Mentors wrt Mentors
#finalpairup2 contains Final Matchings for Startups and Mentors wrt Startups
finalpairup = defaultdict(lambda:[])

for i in l1:
    for j in range(len(i)-2):
        i[j] = int(i[j])
        if i[j]==74:
            continue
        if Mend[str(i[j])][0] > 0:
            Mend[str(i[j])][0] -= 1
            finalpairup[i[-2]].append('M'+str(i[j]))
            i[-1] += 1
            if i[-1] == 2:
                break
nosessions = defaultdict(lambda:0)
#nosessions stores number of sessions
finalpairup2 = defaultdict(lambda:[])
less = []
for i in range(9,47):
    s1 = 'S'+str(i)
    for i1 in finalpairup[i]:
        nosessions[i1] += 1
        finalpairup2[i1].append(s1)
    print(s1, finalpairup[i])
    if len(finalpairup[i]) < 6:
        less.append(s1)
NoSession = []
for i in range(1,103):
    s1 = 'M'+str(i)
    print(s1, nosessions[s1])
    if nosessions[s1] == 0 and s1 != 'M74':
        NoSession.append(s1)

#Part Given Below makes Output Sheets

import xlwt
wb5 = xlwt.Workbook()
sh1 = wb5.add_sheet('Mentor To Startups')
sh2 = wb5.add_sheet('Startups To Mentors')
sh3 = wb5.add_sheet('Startup Pref 2 Slots')
i1 = i2 = 1
while i1 < 103:
    j1 = 0
    for j in range(len(finalpairup2['M'+str(i1)])):
        sh1.write(i2,0,'M'+str(i1)+'X'+str(j+1))
        sh1.write(i2,1,'M'+str(i1))
        sh1.write(i2,2,str(j+1))
        sh1.write(i2,3,finalpairup2['M'+str(i1)][j])
        j1 += 1
        i2+=1
    while j1 < 6:
        sh1.write(i2,0,'M'+str(i1)+'X'+str(j1+1))
        sh1.write(i2,1,'M'+str(i1))
        sh1.write(i2,2,str(j1+1))
        sh1.write(i2,3,'')
        j1 += 1
        i2+=1
    i1 += 1
for i in range(9,47):
    sh2.write(i,0,'S'+str(i))
    for j in range(len(finalpairup[i])):
        sh2.write(i,j+1,finalpairup[i][j])
for i in range(1,103):
    s1 = 'M'+str(i)
    if i == 74:
        continue
    sh3.write(i,0,s1)
    sh3.write(i,1,nosessions[s1])
    if nosessions[s1] == 0 and s1 != 'M74':
        NoSession.append(s1)
wb5.save('Top30PrefSlots01.xlsx')




