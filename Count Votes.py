from xlrd import open_workbook

sheet = open_workbook('MatchedVIDs.xls').sheet_by_index(0)
candidate={
    "Vedang":{"Votes":0,"Approval Rating":0},"Purvai":{"Votes":0,"Approval Rating":0},
    "Vipasha":{"Votes":0,"Approval Rating":0},"Ashutosh":{"Votes":0,"Approval Rating":0},
    "Aryan":{"Votes":0,"Approval Rating":0},"Devak":{"Votes":0,"Approval Rating":0}
}
c_list=list(candidate)
c_name=''
for i in range(sheet.nrows):
    if i == 0: continue
    vote=sheet.row_values(i)
    for j,val in enumerate(vote):
        if j == 0: continue
        if val in c_list:
            candidate[val]["Votes"]+=1
            c_ind=c_list.index(val)
            if c_ind%2 == 0:
                ind= c_ind + 1
            else:
                ind = c_ind - 1
            c_name=c_list[ind]
        else:
            candidate[c_name]["Approval Rating"]+=val

for n,val in candidate.items():
    v=int(val.get("Votes"))
    A=int(val.get("Approval Rating"))
    val["Average Approval Rating of Non Voters"]=round(A/(31-v),1)
    app = A - 2*(31-v)
    val["Overall Approval Rating"]= round((app*0.1 + v)/31.0,2)
    val["Percent of Vote"] = str(round((v/31.0) * 100.0, 1)) + '%'

for cname , prop in candidate.items():
    print(cname)
    for vote, value in prop.items():
        print ("  ",vote,value)

# #################################################
#
# OUTPUT
#
# Vedang
#    Votes 19
#    Approval Rating 81.0
#    Average Approval Rating of Non Voters 6.8
#    Overall Approval Rating 0.8
#    Percent of Vote 61.3%
# Purvai
#    Votes 12
#    Approval Rating 118.0
#    Average Approval Rating of Non Voters 6.2
#    Overall Approval Rating 0.65
#    Percent of Vote 38.7%
# Vipasha
#    Votes 20
#    Approval Rating 65.0
#    Average Approval Rating of Non Voters 5.9
#    Overall Approval Rating 0.78
#    Percent of Vote 64.5%
# Ashutosh
#    Votes 11
#    Approval Rating 116.0
#    Average Approval Rating of Non Voters 5.8
#    Overall Approval Rating 0.6
#    Percent of Vote 35.5%
# Aryan
#    Votes 19
#    Approval Rating 89.0
#    Average Approval Rating of Non Voters 7.4
#    Overall Approval Rating 0.82
#    Percent of Vote 61.3%
# Devak
#    Votes 12
#    Approval Rating 130.0
#    Average Approval Rating of Non Voters 6.8
#    Overall Approval Rating 0.68
#    Percent of Vote 38.7%