## TO DO/NOTES ##

# May want to rewrite for Python 2 so tool can be used in ArcMap?
# Make compatible output excel format for use in aim.analysis
# Including indicator metadata - what do they mean?
# add boxplots from matplotlib for output reports
# add benchmark group to reporting unit summary tab but add additional excel tab for reporting unit summary which combines benchmark groups
# improve bar plot coloring and x axis labels

# select certain indicators based on specific workflows? - suggestion from Bill -

#Import Modules
import arcpy
##import arcgis
##import datetime
##from shutil import copyfile
import os
import sys
#import getpass
import xlrd, xlwt # I think we should be using openpyxl to create xlsx
import openpyxl
from openpyxl.utils import get_column_letter # can only run openpyxl in arcpro
from openpyxl.styles import Font
import pandas as pd
from openpyxl.chart import BarChart, Series, Reference
import matplotlib.pyplot as plt
##import subprocess
from itertools import islice

# import seaborn from local dir
# this path is hardcoded for now, will need to update when the toolbox gets moved to AIMDataTools
libraryPath = r'\\blm\dfs\loc\EGIS\ProjectsNational\AIM\Data\Development\Reports\Benchmark\Terrestrial Benchmark Tool v4.0\seaborn'
sys.path.append(libraryPath)
import seaborn as sns

########## PARAMETER INPUTS ##############

AnalysisPoints = arcpy.GetParameterAsText(0) # this should be a points from step 2

Use_Default_Benchmarks = arcpy.GetParameterAsText(1)

State = arcpy.GetParameterAsText(2)

configFile = arcpy.GetParameterAsText(3) # benchmark config file aka your monitoring objectives

if State == 'National':
    configFile = r"\\blm\dfs\loc\EGIS\ProjectsNational\AIM\Data\Development\Reports\Benchmark\Terrestrial Benchmark Tool v4.0\Monitoring Objectives National.xlsx"

# Potential feature additions:
# Include climate info? REST service form climate engine could be an option
# include trend analysis tab for ESR?
# Include summary stats per reporting unit

########## OUTPUT PATHS ##################
###### define these based on input fc
# we want to just save in the same directory as the point feature class. this should be a gdb for the feature class but a folder for .
desc = arcpy.Describe(AnalysisPoints)
outGDB = desc.path
folder = os.path.dirname(outGDB)
outExcel = folder + "\\" + "benchmarked_points.xlsx" # this excel should eventually have colnames which match up with aim.analysis inputs/benchmark tool
OutputBase = outGDB + "\\" + "benchmarked"
outFC = "Benchmarked_points"

#Create output environment but only if it doesnt already exist
if not os.path.exists(outGDB):
    arcpy.CreateFileGDB_management(os.path.dirname(outGDB), os.path.basename(outGDB))

# BM stats array
bmStats = []
bmColumns = ["Objective" , "Benchmark Group", "Condition Rating", "Total Plots", "Total Plots in Group", "Number of Plots", "Percent of Plots"]

# Open the workbook
xl_workbook = xlrd.open_workbook(configFile) # may want to copy this config file to the output excel?
# Assume in the first sheet
xl_sheet = xl_workbook.sheet_by_index(0)

# Field names
row = xl_sheet.row(0)

# Doing all this in_memory
bmOut = "in_memory/BMOut"

arcpy.FeatureClassToFeatureClass_conversion(AnalysisPoints, "in_memory", "BMOut")

totalPlots = int(arcpy.GetCount_management(bmOut)[0])

# perhaps pull this into a data frame for easier calcs...
# for lotic - benchmark group depends on indicator
# hard code for those 3 indicators
# get lotic indicators into config format
## for benchmark groups - there are defaults but these can be overwritten
for rowNum in range(1, xl_sheet.nrows):
    line = xl_sheet.row(rowNum)

    meetingPlots = 0
    groupQuery = "BenchmarkGroup" + '=' +  "'" + line[3].value + "'"
    group = line[3].value
    indicator = line[7].value
    lowerRelation = str(line[6].value)
    lowerLimit = str(line[5].value)
    upperRelation =  str(line[8].value)
    upperLimit = str(line[9].value)
    rule = indicator + " " + lowerRelation + lowerLimit + " " + 'and' + " " + indicator + " " + upperRelation + upperLimit
    outField = line[0].value
    outAlias = line[1].value
    metMessage = "'" + line[11].value + "'"

    # Add the field if it doesnt already exisit and populate it using the query and message
    if arcpy.ListFields(bmOut, outField): #if field exists, evaluates to true
           arcpy.AddMessage("Applying benchmark " + str(rowNum))
    else:
        arcpy.AddField_management(bmOut, outField, "TEXT", field_length=255, field_alias=outAlias)
        arcpy.AddMessage("Applying benchmark " + str(rowNum))

    # Make the view
    arcpy.MakeTableView_management(bmOut, "memView")

    # Run the group query first to get a count for the total
    if groupQuery == "BenchmarkGroup='All'":
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

#removed notmetmessage here
    arcpy.Delete_management("memView")

    # Add the calcs to list
    bmStats.append(outAlias + "," + group + ","  + metMessage  + "," +  str(totalPlots)  + "," +  str(totalGroupPlots)  + "," + str(meetingPlots)  + "," +  str(percentMeeting))

# Export to Excel
if os.path.exists(outExcel):
    try:
        os.remove(outExcel)
    except:
        arcpy.AddError("Cannot delete excel output. Please close excel files")

arcpy.TableToExcel_conversion(bmOut, outExcel, "ALIAS")

# Export to GDB
try:
    arcpy.FeatureClassToFeatureClass_conversion(bmOut, outGDB, outFC)
except:
    arcpy.AddError("Failed to create new output feature class, check to see if Benchmarked_points feature class already exists")

# Write stats to excel
wb = openpyxl.load_workbook(outExcel)
ws_data = wb.active

# renaming existing sheet
ws_data.title = "Raw Data"

# pull raw data into pandas dataframe, filter to relevant indicators and add conditional formatting for meeting/not meeting
data = ws_data.values
cols = next(data)[1:]
data = list(data)
data = (islice(r, 1, None) for r in data)
df_rawdata = pd.DataFrame(data, columns=cols)

# Select only relevant columns (Indicators) from raw data
# first grab the unique indicators specified in the config file
indicators = xl_sheet.col_values(7)[1:] # removing the first row since thats the header
indicators = list(set(indicators))

# also want the benchmark group column
summary_cols = indicators.copy() # giving this a different name since we want to preserve the list of indicators for plotting later
summary_cols.append("Benchmark Group")
summary_cols.append("PrimaryKey")
df_ind_summary = df_rawdata[summary_cols]

# Group by benchmark group

# Summarise (mean, standard error, sample size)
# this wont work well if theres no benchmark groups defined
df_ind_summary = df_ind_summary.groupby(by ="Benchmark Group", dropna=False).describe(percentiles = []).applymap(lambda x: f"{x:0.2f}") # this defaults to only summarising the numeric columns
# Im also reducing to 3 significant digit3

# copy indicator summary to excel sheet
# create new sheet
wb.create_sheet("Indicator Summary")

writer = pd.ExcelWriter(outExcel, engine = 'openpyxl', mode = 'a')
writer.wb = wb

# need to tell excelwriter what sheets already exist
writer.sheets = dict((ws.title, ws) for ws in wb.worksheets)
df_ind_summary.to_excel(writer, sheet_name = "Indicator Summary", index = True)
writer.save()
writer.close()

# adjust formatting of indicator summary sheet
# remove blank cells under headers by merging
# A2:A3 cells are always the same so can hard code
ws_ind_summary = wb["Indicator Summary"]

# we first have to move the header text since openpyxl will only keep the top left cell
ws_ind_summary.move_range('A3', rows = -1)

# then we can just delete the 3rd row
ws_ind_summary.delete_rows(3)

# looping through each relevant col and row to format each numeric cell
# format numbers in cols C:E
for col in range(2, ws_ind_summary.max_column+1):
    for row in range(3, ws_ind_summary.max_row+1):
        ws_ind_summary[get_column_letter(col)+ str(row)] = float(ws_ind_summary[get_column_letter(col)+ str(row)].value)

# Creating sheet for Benchmarks summary
ws = wb.create_sheet("Reporting Unit Summary")

# Add column headers
ws.append(bmColumns)

# Add in elements of bmStats to each row
for line in bmStats:
    statsLine = line.split(",")
    ws.append(statsLine)

# change table formatting
# bold headers
for col in range(1, len(bmColumns)+1):
    ws[get_column_letter(col) + '1'].font = Font(bold = True)

# looping through each relevant col and row to format each numeric cell
# format numbers in cols C:E
for col in range(4,7):
    for row in range(2, ws.max_row+1):
        ws[get_column_letter(col)+ str(row)] = int(ws[get_column_letter(col)+ str(row)].value)

# format percentage in col F
for row in range(2, ws.max_row+1):
    ws['G'+ str(row)] = float(ws['G'+ str(row)].value)

# remove single quotes from third col
for row in range(2, ws.max_row+1):
    cell = ws['C'+ str(row)].value
    ws['C'+ str(row)].value = cell.replace("'", "")

# can also use seaborn
# although this needs an install...got put the seaborn package in the directory
##
##for i in indicators:
##    sns.set_theme()
##    sns.boxplot(x = 'Benchmark Group', y = i)
##
##    plt.savefig(folder + "/" + i + '_boxplot.png')

# ideally we should also plot the benchmark line on here
# may want to also just copy this graph to the excel sheet

# first grab the lsit of unique objectives from config file
objectives = xl_sheet.col_values(1)[1:] # removing the first row since thats the header
objectives = list(set(objectives))

# we'll just use the raw data dataframe
# if we want all objectives in same plot well need to pivot the df

##for objective in objectives:
##    plt.clf()
##    data = df_rawdata[objective]
##    data.value_counts().plot(kind = 'bar')
##    plt.savefig(folder + "/" + objective + '_histogram.png')
# need to reorder categories some how...

##if os.path.exists(outExcel):
  #  os.remove(outExcel)

# Add plot summary tab
# this should have plot id/primarykey, indicator value, benchmark group, benchmark category

# select columns from raw data
plot_cols = ["PrimaryKey", "PlotID", "Benchmark Group"]
plot_cols.extend(indicators)
# also need to add objective category columns to this list
plot_cols.extend(objectives)

df_plotsummary = df_rawdata[plot_cols]

# create new sheet
wb.create_sheet("Plot Summary")

writer2 = pd.ExcelWriter(outExcel, engine = 'openpyxl', mode = 'a')
writer2.wb = wb

# need to tell excelwriter what sheets already exist
writer2.sheets = dict((ws.title, ws) for ws in wb.worksheets)

# write plot summary df to excel
df_plotsummary.to_excel(writer2, sheet_name = "Plot Summary", index = False)
writer2.save()
writer2.close()

# need to widen cols in all sheets
for sheet in wb.worksheets:
    dims = {}
    for row in sheet.rows:
        for cell in row:
            if cell.value:
                dims[cell.column_letter] = max((dims.get(cell.column_letter, 0), len(str(cell.value))))
    for col, value in dims.items():
        sheet.column_dimensions[col].width = value

wb.save(outExcel)

# Adding Outputs to current project
aprx = arcpy.mp.ArcGISProject("CURRENT")
aprxMap = aprx.activeMap
results = outGDB + "\\" + outFC
aprxMap.addDataFromPath(results)
