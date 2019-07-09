import boto3
import pandas as pd
import datetime
import xlwt
from datetime import timedelta
from xlwt import Workbook
import timedelta
import os
import pytz

#Change ONLY the below variables, role and region. Look below for the specific wording used for each role. Keep both
#variables within quotation marks.
#"cs-prod" "cs-test" "fintech" "venuetech" "rome" "it-support" "data-science" "data-bricks"
role = "data-science"
region = "us-east-1"

os.environ['AWS_PROFILE'] = role
os.environ['AWS_DEFAULT_REGION'] = region
client = boto3.client('workspaces', region_name = region)

outputFile = Workbook()

sheet1 = outputFile.add_sheet(role + " " + region)

#Add each category to the excel output file.
sheet1.write(1, 0, "Username")
sheet1.write(1, 1, "Workspace ID")
sheet1.write(1, 2, "Compute")
sheet1.write(1, 3, "Running Mode")
sheet1.write(1, 4, "Root Volume")
sheet1.write(1, 5, "User Volume")
sheet1.write(1, 6, "Status")
sheet1.write(1, 7, "Notes - Retain/Terminate")
sheet1.write(1, 8, "User Last Active")
sheet1.write(1, 9, "Status")
sheet1.write(1, 10, "Cost Savings")



#Try to pull email addresses if possible.

#Create a list of WorkSpace IDs

#Basic Overview:

#Needs arguments?
dictResponse = client.describe_workspaces()

workSpaceResponseList = dictResponse.get("Workspaces")
listOfWorkSpaces = []

#Gets each WorkspaceId and appends it to listOfWorkSpaces
for eachDict in workSpaceResponseList:
    listOfWorkSpaces.append(eachDict.get("WorkspaceId"))

listOfOldWorkspaces = []
dictOfWorkspacesAndTime = {}

utc = pytz.UTC
dateTimeNow = utc.localize(datetime.datetime.utcnow())

#Go through each workspace and if they are older than 90 days or the information is unavailable,
#adds them to listOfWorkSpaces. Also adds workspace and last entry time to a dictionary to be accessed later.
for eachWorkspace in listOfWorkSpaces:
    workSpaceInformation = client.describe_workspaces_connection_status(WorkspaceIds=[eachWorkspace])
    workSpaceInfoDict = client.describe_workspaces(WorkspaceIds=[eachWorkspace])
    dateTime = workSpaceInformation.get("WorkspacesConnectionStatus")[0].get("LastKnownUserConnectionTimestamp")
    if dateTime == None or dateTimeNow - dateTime > datetime.timedelta(days=90):
        listOfOldWorkspaces.append(eachWorkspace)
        dictOfWorkspacesAndTime[eachWorkspace] = dateTime

columnTracker = 2

listOfWorkspacesWithInfo = []
for everyWorkspace in listOfOldWorkspaces:
    workSpaceInfoDict = client.describe_workspaces(WorkspaceIds=[everyWorkspace])
    dictOfAllInfo = workSpaceInfoDict["Workspaces"][0]
    sheet1.write(columnTracker, 0, dictOfAllInfo["UserName"])
    sheet1.write(columnTracker, 1, dictOfAllInfo["WorkspaceId"])
    sheet1.write(columnTracker, 2, dictOfAllInfo["WorkspaceProperties"]["ComputeTypeName"])
    sheet1.write(columnTracker, 3, dictOfAllInfo["WorkspaceProperties"]["RunningMode"])
    sheet1.write(columnTracker, 4, dictOfAllInfo["WorkspaceProperties"]["RootVolumeSizeGib"])
    sheet1.write(columnTracker, 5, dictOfAllInfo["WorkspaceProperties"]["UserVolumeSizeGib"])
    sheet1.write(columnTracker, 6, dictOfAllInfo["State"])
    sheet1.write(columnTracker, 8, str(dictOfWorkspacesAndTime[everyWorkspace]))
    columnTracker += 1


outputFile.save("WorkspaceInfo.xls")
print("Finished!")
