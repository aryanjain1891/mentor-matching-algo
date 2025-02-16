import xlrd
xlrd.xlsx.ensure_elementtree_imported(False, None)
xlrd.xlsx.Element_has_iter = True
from collections import defaultdict
loc = ("//Users//abirabh//Documents//Top30Real02Mentors.xlsx")
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)
rows = sheet.nrows
cols = sheet.ncols
Mend = {}
Startd = {}
for i in range(1,rows): #Taking the Data Out of the Mentor Expertise Sheet to Mend
    l = []
    for j in range(1,cols):
        x2 = sheet.cell_value(i, j)
        if type(x2)==float:
            l.append(str(int(x2)))
        else:
            l.append(x2)
    l[-1] = int(l[-1])
    Mend[sheet.cell_value(i, 0)] = l
loc2 = ("//Users//abirabh//Documents//Top30Real01Startups.xlsx") 
wb = xlrd.open_workbook(loc2)
sheet2 = wb.sheet_by_index(0)
rows2 = sheet2.nrows
cols2 = sheet2.ncols
for i in range(1,rows2): #Taking the Data Out of the Startup's Mentorship Sector Preferences to Startd
    l = []
    for j in range(1,cols2):
        x2 = sheet2.cell_value(i, j)
        if type(x2) == float:
            l.append(str(int(x2)))
        else:
            l.append(x2)
    l.append(0)
    Startd[str(int(sheet2.cell_value(i, 0)))] = l
pairing = {}
#Matching Algorithm, Takes Startup's Preferences and Overlaps with Each Mentor's Expertise, Provides best overlap
#Have Provided Higher Weightage to Domain of Mentor and Domain of Mentorship, and smaller weightage to specifics
for j in Mend:
    tempscore = []
    for k21 in Startd:
        l00 = Startd[k21]
        score = 0
        l01 = Mend[j]
        k00 = l00[0].split(',')
        k01 = l00[1].split(',')
        k02 = l00[2].split(',')
        k03 = l01[0].split(',')
        k04 = l01[1].split(',')
        k05 = l01[2].split(',')
        for k1 in k00:
            for k2 in k03:
                if k1 == '' or k2 == '':
                    continue
                if int(k1) == int(k2):
                    score += 100
        for k1 in k01:
            for k2 in k04:
                if k1 == '' or k2 == '':
                    continue
                if int(k1) == int(k2):
                    score += 100
        for k1 in k02:
            for k2 in k05:
                if k1 == '' or k2 == '':
                    continue
                if int(k1) == int(k2):
                    score += 200
        if score >= 200:
            tempscore.append((score, k21)) 
    tempscore.sort(reverse = True)
    pairing[j] = tempscore
final = {}
prio = []
for i in Mend:
    prio.append((Mend[i][3], i))
prio.sort(reverse = True)
pairup = []
nosessions2 = defaultdict(lambda:0)
for i in pairing:
    for j in pairing[i]:
        pairup.append([j[0],(i,j[1])])
pairup.sort(reverse = True)
for j in pairup:
    nosessions2[j[1][1]] += 1
finalpairup = defaultdict(lambda:[])
nosessions = defaultdict(lambda:0)
sessions = 0
#Pairings contain all possible mentor-startup pairings with a score greater than 200, now we will start alloting 4 sessions per mentor
for j in pairup:
    Mend[j[1][0]][3] = int(Mend[j[1][0]][3])
    if Mend[j[1][0]][3] > 0 and Startd[j[1][1]][-1] < 4:
        Mend[j[1][0]][3] -= 1
        Startd[j[1][1]][-1] += 1
        finalpairup[j[1][0]].append([j[1][1],j[0]])
#finalpairup contains Final Matchings for Startups and Mentors wrt Mentors
#finalpairup2 contains Final Matchings for Startups and Mentors wrt Startups
#nosessions stores number of sessions
Nosessions3 = []
for i in range(1,102):
    s1 = 'M'+str(i)
    print(s1,finalpairup[s1])
    if finalpairup[s1] == []:
        Nosessions3.append(s1)
    for k in range(len(finalpairup[s1])):
        nosessions[finalpairup[s1][k][0]] += 1
        sessions += 1
finalpairup2 = defaultdict(lambda:[])
for j in finalpairup:
    for k in range(len(finalpairup[j])):
        finalpairup2[('S'+finalpairup[j][k][0])].append(j)

#Part below checks which mentor has how many sessions.

less = []
sessions00 = sessions01 = 0
for j in range(1,47):
    s1 = 'S'+str(j)
    print(s1,nosessions[str(j)])
    if nosessions[str(j)] < 4:
        less.append(s1)
        sessions01 += nosessions[str(j)]
    else:
        sessions00 += nosessions[str(j)]

#Part Given Below makes Output Sheets

print(Nosessions3)
print(less)
print(sessions00, sessions01)
print('DUNZO') 

import xlwt
wb5 = xlwt.Workbook()
sh1 = wb5.add_sheet('Mentor To Startups')
sh2 = wb5.add_sheet('Startups To Mentors')
for i in range(1,103):
    sh1.write(2*i-1,0,'M'+str(i))
    for j in range(len(finalpairup['M'+str(i)])):
        sh1.write(2*i-1,j+1,'S'+finalpairup['M'+str(i)][j][0] + '  '+str(finalpairup['M'+str(i)][j][1]))
        sh1.write(2*i,j+1,'M'+str(i)+'_'+str(j+1))
for i in range(1,47):
    sh2.write(i,0,'S'+str(i))
    for j in range(len(finalpairup2['S'+str(i)])):
        sh2.write(i,j+1,finalpairup2['S'+str(i)][j])
wb5.save('Top30MatchingRealFinal12.xlsx')


