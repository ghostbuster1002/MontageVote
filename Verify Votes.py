from xlrd import open_workbook
from xlwt import Workbook

votedsheet = open_workbook('Montage Voting Results.xlsx').sheet_by_index(2)
UID=[x.value for x in votedsheet.col(0)]
UIDsheet=open_workbook('Vids.xls').sheet_by_index(0)
VID=[x.value for x in UIDsheet.col(0)]
ind=[0]
for id in VID:
    try:
        ind.append(UID.index(id))
    except:
        continue
wb = Workbook()
sheet = wb.add_sheet('Sheet 1')
for i,index in enumerate(ind):
    vote=votedsheet.row_values(index)
    for x, val in enumerate(vote):
        sheet.write(i, x, val)
    wb.save('MatchedVIDs.xls')