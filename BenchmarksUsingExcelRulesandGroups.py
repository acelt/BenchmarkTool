import xlrd, arcpy, os, xlwt

#inTerra = r"C:\tmp\Doug\Export\AIMExport1-30-20.gdb\TerrADat"
inTerra = r"C:\Users\alaurencetraynor\Documents\2021\DesertTortoise\DesertTortoise.gdb\LMF_tortoise"

configFile = r"C:\Users\alaurencetraynor\Documents\2021\Tortoise.xlsx"

outExcel = r"C:\Users\alaurencetraynor\Documents\2021\DesertTortoise\TortoiseData.xls"
outStatsExcel = r"C:\Users\alaurencetraynor\Documents\2021\DesertTortoise\Tortoisetats.xls"
outGDB = r"C:\Users\alaurencetraynor\Documents\2021\DesertTortoise\TortoiseData.gdb"
outFC = "TortoiseData"

# BM stats array
bmStats = []
bmColumns = ["Benchmark" , "Condition Rating", "Total Plots", "Total Plots in Group", "Number of Plots Meeting Benchmark", "Percent of Plots Meeting Benchmark"]


# Open the workbook
xl_workbook = xlrd.open_workbook(configFile)
# Assume in the first sheet
xl_sheet = xl_workbook.sheet_by_index(0)

# Field names
row = xl_sheet.row(0)


# Doing all this in_memory
bmOut = "in_memory/BMOut"
arcpy.FeatureClassToFeatureClass_conversion(inTerra, "in_memory", "BMOut")

totalPlots = int(arcpy.GetCount_management(bmOut)[0])


for rowNum in range(1, xl_sheet.nrows):
    line = xl_sheet.row(rowNum)

    meetingPlots = 0

    groupQuery = line[0].value
    rule = line[1].value
    outField = line[2].value
    outAlias = line[3].value
    metMessage = "'" + line[4].value + "'"
    notmetMessage = "'" + line[5].value + "'"

    # Add the field and populate it using the query and message
    arcpy.AddField_management(bmOut, outField, "TEXT", field_length=255, field_alias=outAlias)

    # Make the view
    arcpy.MakeTableView_management(bmOut, "memView")

    # Run the group query first to get a count for the total
    if groupQuery == "All":
        arcpy.SelectLayerByAttribute_management("memView", "SWITCH_SELECTION")
    else:
        arcpy.SelectLayerByAttribute_management("memView", "NEW_SELECTION", groupQuery)

    totalGroupPlots = int(arcpy.GetCount_management("memView")[0])


    # Then do the rule
    arcpy.SelectLayerByAttribute_management("memView", "SUBSET_SELECTION", rule)
    arcpy.CalculateField_management("memView", outField, metMessage, "PYTHON_9.3")
    meetingPlots = int(arcpy.GetCount_management("memView")[0])
    if totalGroupPlots == 0:
        percentMeeting = 0
    else:
        percentMeeting = round((float(meetingPlots) / totalGroupPlots) * 100,2)

    arcpy.SelectLayerByAttribute_management("memView", "SWITCH_SELECTION")
    arcpy.CalculateField_management("memView", outField, notmetMessage, "PYTHON_9.3")

    arcpy.Delete_management("memView")

    # Add the calcs to list
    bmStats.append(outAlias + "," + metMessage + "," + str(totalPlots) + "," + str(totalGroupPlots) + "," + str(meetingPlots) + "," + str(percentMeeting))


# Export to Excel
if os.path.exists(outExcel):
    os.remove(outExcel)
arcpy.TableToExcel_conversion(bmOut, outExcel, "ALIAS")

# Export to GDB
if os.path.exists(outGDB):
    arcpy.Delete_management(outGDB)
arcpy.CreateFileGDB_management(os.path.dirname(outGDB), os.path.basename(outGDB))
arcpy.FeatureClassToFeatureClass_conversion(bmOut, outGDB, outFC)


# Write stats to excel
wb = xlwt.Workbook()
ws = wb.add_sheet('Area Summary')

styleInt = xlwt.easyxf(num_format_str='0')
styleFloat = xlwt.easyxf(num_format_str='0.00')

# Write out the field names
colNum = 0
for col in bmColumns:
    ws.write(0, colNum, col)
    ws.col(colNum).width = 25 * 367
    colNum += 1

onRow = 1
for line in bmStats:
    statsLine = line.split(",")
    ws.write(onRow, 0, statsLine[0])
    ws.write(onRow, 1, statsLine[1].strip("'"))
    ws.write(onRow, 2, int(statsLine[2]), styleInt)
    ws.write(onRow, 3, int(statsLine[3]), styleInt)
    ws.write(onRow, 4, int(statsLine[4]), styleInt)
    ws.write(onRow, 5, float(statsLine[5]), styleFloat)

    onRow += 1

if os.path.exists(outStatsExcel):
    os.remove(outStatsExcel)
wb.save(outStatsExcel)
