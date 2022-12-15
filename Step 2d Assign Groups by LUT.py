#-------------------------------------------------------------------------------
# Name:       Step 2d
# Purpose:      Add Benchmark groupsby look up table
#
# Author:      alaurencetraynor
#
# Created:     08/03/2022
# Copyright:   (c) alaurencetraynor 2022
# Licence:     <your licence>
#-------------------------------------------------------------------------------

# Import arcpy module
import arcpy
import pandas as pd
import numpy as np

# Script arguments
#Points_Feature_Class = arcpy.GetParameterAsText(0)
Points_Feature_Class = r"\\blm\dfs\loc\EGIS\ProjectsNational\AIM\Data\Development\Reports\Benchmark\greatercurlewspringtest.gdb\greatercurlewspringtest_AIM_points"

# this will need primarykey and a group field
Look_up_Table = arcpy.GetParameterAsText(1)


# from the table columns
Group_Field = arcpy.GetParameterAsText(2)

# import csv
df = pd.read_csv(Look_up_Table, usecols = ['PrimaryKey', Group_Field])
df.to_numpy()

# table join between Primary keys


# Calculate field for benchmark group from group field

# keep only the neccesary fields from the LUT and points feature class

# save output

# add to arcpro project
# Adding Outputs to current project
aprx = arcpy.mp.ArcGISProject("CURRENT")
aprxMap = aprx.activeMap
aprxMap.addDataFromPath()
