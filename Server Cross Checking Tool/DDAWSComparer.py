import boto3
import os
from xlwt import Workbook
import pandas as pd
from datadog import initialize, api
import datadog
import requests
import json
import csv

# Use the following command in the active directory (powershell) to pull server information and save it to a
# csv file in your Desktop under the name ServerNames.csv:
#
# Get-ADComputer -Filter * -Properties samaccountname | Export-Csv -Path .\Desktop\ServerNames.csv"

# PLEASE READ BEFORE RUNNING:
#
# Replace APIKey and APPKey with the appropriate information from DataDog.
# Use this link to access that information: https://app.datadoghq.com/account/settings#api
#
# Also replace the filePath variable with the file path you want the final output data to be saved to.
#
# Replace the CSInventoryFilePath with the file path to the CS Inventory excel spreadsheet.
#
# Replace the ADServerInfoFilePath with the file path to the Active Directory csv file. MAKE SURE to delete the
# topmost row of the Active Directory server file before inputting the file path, as that row has nonessential
# information and will break the script.
#
# Make sure the filePath ends with insert_new_file_name.xls . Use the filePath I have described as a template.
#
# Finally, make sure that the list of regions and profiles is updated. To do this, go over to your config file in
# the .aws file and check to see if each profile and region is in the variable "profiles" and "regions" respectively.
#
# If you have any questions, contact me, Kamila Wickramarachchi, on Slack or through email at
# kamilawickramarachchi@livenation.com.
#
# This code was last updated on 7/29/2019.

# Fill in the variables below. Many of them have prefilled values to show the format for the variables. Do NOT use the
# prefilled values for APIKey, APPKey, filePath, CSInventoryFilePath, and ADServerInfoFilePath.
# Feel free to change the values in regions/profiles or leave them alone.

APIKey = ""
APPKey = ""
filePath = "/Users/kamila.wickramarachchi/Desktop/Scripts/DataDogAWS/ServerInfo.xls"
CSInventoryFilePath = "/Users/kamila.wickramarachchi/Desktop/Scripts/DataDogAWS/CSinventory(Updated).xlsm"
ADServerInfoFilePath = "/Users/kamila.wickramarachchi/Desktop/Scripts/DataDogAWS/ServerNames.csv"

regions = ['eu-north-1', 'ap-south-1', 'eu-west-3', 'eu-west-2', 'eu-west-1', 'ap-northeast-2', 'ap-northeast-1', 'sa-east-1', 'ca-central-1', 'ap-southeast-1', 'ap-southeast-2', 'eu-central-1', 'us-east-1', 'us-east-2', 'us-west-1', 'us-west-2']

profiles = ['cs-test', 'cs-prod', 'fintech', 'venuetech', 'rome', 'infosec-na', 'infosec-sandbox', 'artist-service', 'on-tour-sites', 'hobe', 'it-support', 'network-users', 'data-science', 'data-bricks', 'lninfra-prod', 'microflex']

# The below method pulls the requested AWS Server information for a specific profile and region.

def AWSServersandMetadata(profile, region):

    ec2 = boto3.session.Session(profile_name=profile).client('ec2', region_name=region)

    svrInfo = ec2.describe_instances()

    listOfAllInstancesInfo = []

    for eachReservation in svrInfo["Reservations"]:
        for eachInstance in eachReservation["Instances"]:
            # This save the value of the Dictionary key 'InstanceId'
            listOfAllInstancesInfo.append(eachInstance)

    return listOfAllInstancesInfo

# This method combines the AWS Server information and the DD Server information and generates an excel sheet.
def ExcelFiller(listOfInstanceInfo, infra_content, hostList, profile, region, fileName, CSInventoryFilePath,
                ADServerInfoFilePath, masterInfra_Content, masterHostList):

    # Exits the method and returns false if there is no AWS information for the specified profile and region.
    if len(listOfInstanceInfo) == 0:
        print(profile + " + " + region + " does not have any servers.")
        return False, masterInfra_Content, masterHostList

    os.environ['AWS_PROFILE'] = profile

    ec2 = boto3.client('ec2', region_name=region)

    outputFile = Workbook()

    sheet1 = outputFile.add_sheet("sheet1")

    # Writes the column names
    sheet1.write(0, 0, "Host Name")
    sheet1.write(0, 1, "AWS ID")
    sheet1.write(0, 2, "IP Address")
    sheet1.write(0, 3, "AWS State")
    sheet1.write(0, 4, "DataDog State")
    sheet1.write(0, 5, "In AWS and Datadog?")
    sheet1.write(0, 6, "DataDog Agent")
    sheet1.write(0, 7, "DataDog Agent Type")
    sheet1.write(0, 8, "Exists in AD?")
    sheet1.write(0, 9, "Exists in AnnMarie Excel Doc?")
    sheet1.write(0, 10, "AWS Tag: Account")
    sheet1.write(0, 11, "AWS Tag: Application")
    sheet1.write(0, 12, "AWS Tag: Environment")
    sheet1.write(0, 13, "AWS Tag: Region")
    sheet1.write(0, 14, "In AWS?")
    sheet1.write(0, 15, "In DataDog?")
    sheet1.write(0, 16, "AWS Tag: Name (If not same as Host Name)")

    listInstanceID = []

    file = pd.read_excel(CSInventoryFilePath, sheet_name="AWS")

    # Creates a list that contains all the Instance IDs from the AWS input information.
    for i in file.index:
        listInstanceID.append(file['Instance ID'][i])

    rowTracker = 1

    listSamAccountName = []
    listSamAccountNameAltered = []

    file1 = pd.read_csv(ADServerInfoFilePath)

    # Creates a list that contains the Server Account Name as shown in the Active Directory.
    for i in file1.index:
        listSamAccountName.append(file1['SamAccountName'][i][:-1])
        listSamAccountNameAltered.append(file1['SamAccountName'][i][:-1] + ".LYV.LIVENATION.COM")

    instanceIDList = {}

    # Adds the AWS information to the excel spreadsheet.
    for eachDict in listOfInstanceInfo:
        instanceIDList[str(eachDict['InstanceId'])] = rowTracker
        sheet1.write(rowTracker, 1, str(eachDict['InstanceId']))
        if (str(eachDict['InstanceId'])) in listInstanceID:
            sheet1.write(rowTracker, 9, "Yes")
        else:
            sheet1.write(rowTracker, 9, "No")
        try:
            sheet1.write(rowTracker, 2, str(eachDict['PrivateIpAddress']))
        except Exception as e:
            x = 1
            # Catches the error that occurs if there is no IP Address information.
        try:
            sheet1.write(rowTracker, 3, str(eachDict['State']['Name']))
        except Exception as e:
            x = 1
            # Catches the error that occurs if there is no State information.
        sheet1.write(rowTracker, 14, "Yes")
        for eachTag in eachDict['Tags']:
            if eachTag['Key'] == "Account Name":
                sheet1.write(rowTracker, 10, str(eachTag['Value']))
            elif eachTag['Key'] == "Application":
                sheet1.write(rowTracker, 11, str(eachTag['Value']))
            elif eachTag['Key'] == "Environment":
                sheet1.write(rowTracker, 12, str(eachTag['Value']))
            elif eachTag['Key'] == 'AZ':
                sheet1.write(rowTracker, 13, str(eachTag['Value']))

        rowTracker += 1

    # Adds the DataDog information to the excel spreadsheet.
    for eachServer in infra_content["rows"]:
        try:
            instanceId = [s for s, s in enumerate(eachServer['tags_by_source']['Amazon Web Services']) if 'instance_id' in s][0][12:]
            row = instanceIDList[instanceId]
            masterInfra_Content["rows"].remove(eachServer)
            try:
                operatingSystem = [s for s, s in enumerate(eachServer['tags_by_source']['Amazon Web Services']) if 'os' in s][0][3:]
                if (operatingSystem == "linux") or (operatingSystem == "windows"):
                    sheet1.write(row, 7, operatingSystem)
            except Exception as e:
                x = 1
                # Catches the error that occurs if there is no Operating System information.
            sheet1.write(row, 0, eachServer['host_name'])
            if 'i-' in str(eachServer['host_name']):
                try:
                    instanceInfo = ec2.describe_instances(InstanceIds=[str(eachServer['host_name'])])
                    tags = instanceInfo["Reservations"][0]["Instances"][0]['Tags']
                    for eachTag in tags:
                        if eachTag['Key'] == 'name' or eachTag['Key'] == "Name":
                            hostName = eachTag['Value']
                            sheet1.write(row, 16, hostName)
                except Exception as e:
                    x = 1
                    # Catches the error that occurs if there is no tag information for the instance.
            else:
                hostName = eachServer['host_name']
            if (hostName.upper() in listSamAccountName) or (hostName.upper() in listSamAccountNameAltered):
                sheet1.write(row, 8, "Yes")
            else:
                sheet1.write(row, 8, "No")
            if (eachServer['up']):
                sheet1.write(row, 4, 'Up')
            else:
                sheet1.write(row, 4, '???')
            sheet1.write(row, 15, "Yes")
        except Exception as e:
            x = 1
            # Catches the error that occurs if the instance ID is not in the DataDog information (infra_content).
        try:
            instanceId = eachServer['name']
            row = instanceIDList[instanceId]
            masterInfra_Content["rows"].remove(eachServer)
            try:
                operatingSystem = [s for s, s in enumerate(eachServer['tags_by_source']['Amazon Web Services']) if 'os' in s][0][3:]
                if (operatingSystem == "linux") or (operatingSystem == "windows"):
                    sheet1.write(row, 7, operatingSystem)
            except Exception as e:
                x = 1
                # Catches the error that occurs if there is no Operating System information.
            sheet1.write(row, 0, eachServer['host_name'])
            if 'i-' in str(eachServer['host_name']):
                try:
                    instanceInfo = ec2.describe_instances(InstanceIds=[str(eachServer['host_name'])])
                    tags = instanceInfo["Reservations"][0]["Instances"][0]['Tags']
                    for eachTag in tags:
                        if eachTag['Key'] == 'name' or eachTag['Key'] == "Name":
                            hostName = eachTag['Value']
                            sheet1.write(row, 16, hostName)
                except Exception as e:
                    x = 1
                    # Catches the error that occurs if there is no tag information for the instance.
            else:
                hostName = eachServer['host_name']
            if (hostName.upper() in listSamAccountName) or (hostName.upper() in listSamAccountNameAltered):
                sheet1.write(row, 8, "Yes")
            else:
                sheet1.write(row, 8, "No")
            if (eachServer['up']):
                sheet1.write(row, 4, 'Up')
            else:
                sheet1.write(row, 4, '???')
            sheet1.write(row, 15, "Yes")
        except Exception as e:
            x = 1
            # Catches the error that occurs if the instance ID is not in the DataDog information (infra_content).
        try:
            instanceId = eachServer['aws_id']
            row = instanceIDList[instanceId]
            masterInfra_Content["rows"].remove(eachServer)
            try:
                operatingSystem = [s for s, s in enumerate(eachServer['tags_by_source']['Amazon Web Services']) if 'os' in s][0][3:]
                if (operatingSystem == "linux") or (operatingSystem == "windows"):
                    sheet1.write(row, 7, operatingSystem)
            except Exception as e:
                x = 1
                # Catches the error that occurs if there is no Operating System information.
            sheet1.write(row, 0, eachServer['host_name'])
            if 'i-' in str(eachServer['host_name']):
                try:
                    instanceInfo = ec2.describe_instances(InstanceIds=[str(eachServer['host_name'])])
                    tags = instanceInfo["Reservations"][0]["Instances"][0]['Tags']
                    for eachTag in tags:
                        if eachTag['Key'] == 'name' or eachTag['Key'] == "Name":
                            hostName = eachTag['Value']
                            sheet1.write(row, 16, hostName)
                except Exception as e:
                    x = 1
                    # Catches the error that occurs if there is no tag information for the instance.
            else:
                hostName = eachServer['host_name']
            if (hostName.upper() in listSamAccountName) or (hostName.upper() in listSamAccountNameAltered):
                sheet1.write(row, 8, "Yes")
            else:
                sheet1.write(row, 8, "No")
            if (eachServer['up']):
                sheet1.write(row, 4, 'Up')
            else:
                sheet1.write(row, 4, '???')
            sheet1.write(row, 15, "Yes")
        except Exception as e:
            x = 1
            # Catches the error that occurs if the instance ID is not in the DataDog information (infra_content).


    numberRows = len(listOfInstanceInfo)

    i = 1

    # Writes the (In DD and AWS) column.
    while (i <= numberRows):
        try:
            sheet1.write(i, 15, "No")
            sheet1.write(i, 5, "No")
        except Exception as e:
            sheet1.write(i, 5, "Yes")

        i += 1

    # Adds DataDog agent information to the spreadsheet.
    for eachHost in hostList:
        try:
            instanceId = [s for s, s in enumerate(eachHost['tags_by_source']['Amazon Web Services']) if 'instance_id' in s][0][12:]
            row = instanceIDList[instanceId]
            masterHostList.remove(eachHost)
            if ('agent' in eachHost['apps']):
                sheet1.write(row, 6, 'Yes')
            else:
                sheet1.write(row, 6, 'No')
        except Exception as e:
            x = 1
            # Catches the error that occurs if the instance ID is not in the AWS instanceIDList or the instance ID
            # does not exist.

    # Saves the output excel file.
    outputFile.save(fileName)

    print(profile + " + " + region + " worked!")

    # Returns true if the excel information is succesfully pulled and saved.
    return True, masterInfra_Content, masterHostList

# This method pulls all server information from Data Dog. The information is saved into two variables:
# infra_content and hostList, both of which are returned.
def DDServersandMetadata(APIKey, APPKey):
    options = {'api_key': APIKey,
               'app_key': APPKey}

    initialize(**options)

    s = requests.session()
    s.params = {
        'api_key': APIKey,
        'application_key': APPKey,
    }
    infra_link = 'https://app.datadoghq.com/reports/v2/overview'
    infra_content = s.request(method='GET', url=infra_link, params=s.params).json()

    numberOfHosts = api.Hosts.totals()['total_active']

    start = 0
    end = 100

    hostList = []

    while (end < numberOfHosts):
        hostList += api.Hosts.search(start=start, end=end)['host_list']
        end += 100
        start += 100

    hostList += api.Hosts.search(start=(end-100), end=numberOfHosts)['host_list']

    return infra_content, hostList

infra_content, hostList = DDServersandMetadata(APIKey, APPKey)
masterInfra_Content = infra_content
masterHostList = hostList

fileNameList = []

# The following code creates the excel sheet for each profile and region in the profiles list and the regions list
# respectively.

for eachProfile in profiles:
    for eachRegion in regions:
        try:
            fileName = eachProfile + "-" + eachRegion
            worked, masterInfra_Content, masterHostList = ExcelFiller(AWSServersandMetadata(eachProfile, eachRegion),
                                                                      infra_content, hostList, eachProfile, eachRegion,
                                                                      fileName + ".xls", CSInventoryFilePath,
                                                                      ADServerInfoFilePath, masterInfra_Content,
                                                                      masterHostList)
            if worked:
                fileNameList.append(fileName)
        except Exception as e:
            # Prints if the requested profile and region don't have any servers in them.
            print(eachProfile + " + " + eachRegion + " does not have any servers.")

finalExcelFile = pd.ExcelWriter(filePath)

# Takes each individual profile/region excel file and combines them into one master excel file.
for eachFile in fileNameList:
    tempFileName = pd.read_excel((eachFile + '.xls'), 'sheet1')
    tempFileName.to_excel(finalExcelFile, sheet_name=eachFile)

finalExcelFile.save()
