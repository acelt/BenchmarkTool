## TO DO/NOTES ##

# May want to rewrite for Python 2 so tool can be used in ArcMap
# Make compatible output excel format for use in aim.analysis
# Including indicator metadata - what do they mean?
# Add symbology defualts/lyr file

#Import Modules
import arcpy
##import arcgis
import datetime
##from shutil import copyfile
import os
import sys
##import getpass
##import xlrd, xlwt

########## PARAMETER INPUTS ##############

AnalysisArea = arcpy.GetParameterAsText(0) # this should be a polygon of your reporting unit

AnalysisName = arcpy.GetParameterAsText(1) # base name of your outputs

IncludeLMF = arcpy.GetParameterAsText(2) # benchmark config file aka your monitoring objectives. Eventually this should pull from the first couple arctoolbox tools

MostRecent = arcpy.GetParameterAsText(3)

DateRange_start = arcpy.GetParameterAsText(4)

DateRange_end = arcpy.GetParameterAsText(5)

OutputFolder = arcpy.GetParameterAsText(6)
# diretory to store all outputs

########## OUTPUT PATHS ##################
###### define these based on output folder
outExcel = OutputFolder + "\\" + AnalysisName + "_benchmarked.xlsx" # this excel should eventually have colnames which match up with aim.analysis inputs/benchmark tool
outStatsExcel = OutputFolder + "\\" + AnalysisName + "_summarystats.xlsx"
outGDB = OutputFolder + "\\" + AnalysisName + ".gdb"
OutputBase = outGDB + "\\" + AnalysisName
outFC = AnalysisName + "_benchmarked"

# set environment
arcpy.env.workspace = outGDB

#Create output environment
if os.path.exists(outGDB):
    arcpy.Delete_management(outGDB)
arcpy.CreateFileGDB_management(os.path.dirname(outGDB), os.path.basename(outGDB))

#Save copies of analysis polygons
arcpy.CopyFeatures_management(AnalysisArea, OutputBase + '_polygon')

############################################################################################
## could also try to connect to the service instead of the SDE

#Connect to SDE
start = datetime.datetime.now()
arcpy.AddMessage("Creating a SDE Connection")

# If directory does not exist, add it
if not os.path.exists(OutputFolder):
        arcpy.AddMessage("Output folder does not exist. Creating output folder in " + OutputFolder)
        os.makedirs(OutputFolder)

## Eventually need to update this to use AIM/EGIS connections files instead creating a connection each time
sdeConn =  r"\\blm.doi.net\dfs\loc\EGIS\ProjectsNational\AIM\AIMDataTools\SDE\AIMTerrestrialPub.sde"
terraFC = sdeConn + '/' + "ilmocAIMTerrestrialPub.ilmocAIMPubDBO.TerrestrialIndicators"

if IncludeLMF == 'true':
    arcpy.FeatureClassToFeatureClass_conversion(terraFC, 'in_memory', 'AIMLMF_All')
else:
    arcpy.FeatureClassToFeatureClass_conversion(terraFC, 'in_memory', 'AIMLMF_All', "ProjectName <> 'LMF'")

# Subsetting to analysis area
selection = arcpy.SelectLayerByLocation_management("in_memory\AIMLMF_All","COMPLETELY_WITHIN",AnalysisArea,"","NEW_SELECTION")

arcpy.CopyFeatures_management(selection, "in_memory\AIMLMF_selection")

### Adding ability to filter by date range ###
if DateRange_start is not 'none' and DateRange_end is not 'none':
    start = datetime.datetime.now()
    where_clause = "DateVisited BETWEEN " + "datetime '"+ DateRange_start + " 00:00:00' AND " + "datetime '" + DateRange_end + " 00:00:00'"

    selection = arcpy.SelectLayerByAttribute_management('in_memory\AIMLMF_selection',"ADD_TO_SELECTION", where_clause)
    arcpy.CopyFeatures_management(selection, "in_memory\AIMLMF_newselection")
    current = datetime.datetime.now()
    arcpy.AddMessage("Filtered points by date range and exported, elapsed time: "+ str(current-start))
else:
    arcpy.AddMessage("Date range not provided")
    arcpy.CopyFeatures_management(selection, "in_memory\AIMLMF_newselection")

current = datetime.datetime.now()
arcpy.AddMessage("Finished subsetting AIM points to Analysis Area and exported to GDB: "+ str(current-start))

########### Filtering to the most recent visit #################

start = datetime.datetime.now()

if MostRecent == 'true':

    plots = {}
    with arcpy.da.SearchCursor('in_memory\AIMLMF_newselection', field_names = ["PlotKey","PrimaryKey","DateVisited"]) as cursor:
        for row in cursor:
            if cursor[0] in plots:
                if plots[cursor[0]][1]<cursor[2]: #Check the date for the visits to this plot. Keep the most recent
                    plots[cursor[0]] = [cursor[1],cursor[2]]
            else:
                plots[cursor[0]] = [cursor[1],cursor[2]]
    keys = []
    for plot in plots.values():
        keys.append(str(plot[0]))

    keys = tuple(keys)

    arcpy.SelectLayerByAttribute_management("in_memory\AIMLMF_newselection", "NEW_SELECTION", 'PrimaryKey IN ' +str(keys))
    arcpy.CopyFeatures_management("in_memory\AIMLMF_newselection", OutputBase + '_AIM_points')
    current = datetime.datetime.now()
    arcpy.AddMessage("Filtered points by most recent visit and exported, elapsed time: "+ str(current-start))

    # Add benchmark group field
    arcpy.management.AddField(OutputBase + '_AIM_points', "BenchmarkGroup", "TEXT", "","","100","Benchmark Group")

    # Add Reporting unit field
    arcpy.management.AddField(OutputBase + '_AIM_points', "ReportingUnit", "TEXT", "","","100","Reporting Unit")

    #Todo : Add check if the selection results in an empty layer so that things can fail gracefully

else:
    arcpy.CopyFeatures_management("in_memory\AIMLMF_newselection", OutputBase + '_AIM_points')

    # Add benchmark group field
    arcpy.management.AddField(OutputBase + '_AIM_points', "BenchmarkGroup", "TEXT", "","","100","Benchmark Group")

    # Add Reporting unit field
    arcpy.management.AddField(OutputBase + '_AIM_points', "ReportingUnit", "TEXT", "","","100","Reporting Unit")

# ideally want a template layer file for these points which aligns with the service

##if TablesCheckbox == 'true':
##
##    if "All Tables" in TablesList:
##        if IncludeLMF == 'true':
##            TablesList = ["CONCERN", "COUNTYNM", "DISTURBANCE", "ESFSG","GINTERCEPT","GPS","PASTUREHEIGHTS","PINTERCEPT","PLANTCENSUS","POINT","POINTWEIGHT","PRACTICE","PTNOTE","RANGEHEALTH","RHSUMMARY","SOILDISAG","SOILHORIZON","STATENM","tblGapDetail","tblGapHeader","tblLines","tblLPIDetail","tblLPIDetail","tblMissingDataKnownErrors","tblPeople","tblPeople","tblPlantDenHeader","tblPlantDenQuads","tblPlantDenSpecies","tblPlantProdDetail","tblPlantProdHeader","tblPlotNotes","tblPlots","tblQualDetail","tblQualDetail","tblSites","tblSoilPitHorizons","tblSoilPits","tblSoilStabDetail","tblSoilStabHeader","tblSpecies","tblSpecies","tblSpecRichDetail","tblSpecRichHeader","tblStateSpecies","TerrestrialSpecies"]
##        else:
##            TablesList = ["tblGapDetail","tblGapHeader","tblLines","tblLPIDetail","tblLPIDetail","tblMissingDataKnownErrors","tblPeople","tblPeople","tblPlantDenHeader","tblPlantDenQuads","tblPlantDenSpecies","tblPlantProdDetail","tblPlantProdHeader","tblPlotNotes","tblPlots","tblQualDetail","tblQualDetail","tblSites","tblSoilPitHorizons","tblSoilPits","tblSoilStabDetail","tblSoilStabHeader","tblSpecies","tblSpecies","tblSpecRichDetail","tblSpecRichHeader","tblStateSpecies","TerrestrialSpecies"]
##
##    pkList = [row[0] for row in arcpy.da.SearchCursor("in_memory\AIMLMF_selection", ["PrimaryKey"])]
##    stateList = set([row[0] for row in arcpy.da.SearchCursor("in_memory\AIMLMF_selection", ["SpeciesState"])])
##
##    arcpy.AddMessage("Gathering related tables")
##
##    sdePrefix = "ilmocAIMTerrestrialPub.ILMOCAIMPUBDBO."
##
##     # FCs are special case
##    fcList = ["tblPlots", "TerrestrialSpecies"]
##
##    for table in TablesList:
##
##        arcpy.AddMessage("Gathering " + table)
##
##        tablePath = sdeConn + "\\" + sdePrefix + table
##
##        # FCs are special case
##        if table in fcList:
##            arcpy.MakeFeatureLayer_management(tablePath, "tempTV")
##        else:
##            arcpy.MakeTableView_management(tablePath, "tempTV")
##
##        # We want each PrimaryKey
##        for pk in pkList:
##            arcpy.SelectLayerByAttribute_management("tempTV", "NEW_SELECTION", 'PrimaryKey = ' + "'" + pk + "'")
##            selectCount = int(arcpy.GetCount_management("tempTV").getOutput(0))
##            if selectCount > 0:
##                arcpy.AddMessage( "     Found " + str(selectCount) + " records")
##                arcpy.CopyFeatures_management("tempTV", outGDB + "\\" + table)
##            else:
##               arcpy.AddMessage( "     No records found for PrimaryKey " + pk)
##
##            arcpy.SelectLayerByAttribute_management("tempTV", "CLEAR_SELECTION")
##
##        arcpy.Delete_management("tempTV")
##
##        if table == "tblStateSpecies":
##            for speciesState in stateList:
##                arcpy.AddMessage( "\nAdding records from tblStateSpecies for " + speciesState)
##                arcpy.SelectLayerByAttribute_management("tempTV", "NEW_SELECTION", 'SpeciesState = ' + "'" + speciesState + "'")
##                arcpy.CopyFeatures_management("tempTV", outGDB + "\\" + "tblStateSpecies")
##            arcpy.Delete_management("tempTV")

# Adding Outputs to current project
aprx = arcpy.mp.ArcGISProject("CURRENT")
aprxMap = aprx.activeMap
points = OutputBase+'_AIM_points'
polygon = OutputBase+'_polygon'

aprxMap.addDataFromPath(polygon)
aprxMap.addDataFromPath(points)

current = datetime.datetime.now()
arcpy.AddMessage("All done: "+ str(current-start))
