import boto3
from xlwt import Workbook
import pandas as pd


# READ ME:
#
# This program has very few arguments. The only three arguments you should alter are below: Server, profiles, and
# regions. If you want to look at all servers, set Server to True. If you want to look for unattached volumes, set
# Server to False. profiles is just the list of AWS profiles you want the program to go through when looking for
# servers or unattached volumes. regions is just the list of AWS regions you want the program to go through when
# looking for servers or unattached volumes.
#
# The output file is saved as OutputFile.xls. The other excel files that will be saved can be deleted once the program
# is completed. All files will be saved in the directory in which this python script is saved/run.
#
# Be aware that this program may take a long time to run (up to 30 minutes or even more at times).
# To reduce how long the program takes, loop through less profiles and less regions by removing profiles and regions
# from their respective lists.
#
# For servers, the output information is sorted by whether the instance is running or stopped, and is further sorted by
# Business Owner.

Server = False

profiles = ['cs-test', 'cs-prod', 'fintech', 'venuetech', 'rome', 'infosec-na', 'infosec-sandbox', 'artist-service', 'on-tour-sites', 'hobe', 'it-support', 'network-users', 'data-science', 'data-bricks', 'lninfra-prod', 'microflex']

regions = ['us-east-1', 'us-west-2']

# Function below pulls the information for all instances in a specified profile and region.

def instanceInfoPuller(profile, region):

    ec2 = boto3.session.Session(profile_name=profile).client('ec2', region_name=region)

    svrInfo = ec2.describe_instances()

    listOfAllInstancesInfo = []

    for eachReservation in svrInfo["Reservations"]:
        for eachInstance in eachReservation["Instances"]:
            # This will output the value of the Dictionary key 'InstanceId'
            listOfAllInstancesInfo.append(eachInstance)

    return listOfAllInstancesInfo

# Writes Excel File with Categories: Server Name, Instance ID, Instance Type, Associated Volumes Name, Volume ID,
# EBS Volume Size, Storage Type, IOPS

# The function below takes in a list of Instance Info (received from instanceInfoPuller). The function also takes in a
# profile, region, and a fileName to which the output data will be stored.

def createExcel(listOfAllInstancesInfo, profile, region, fileName):

    if len(listOfAllInstancesInfo) == 0:
        return False

    outputFile = Workbook()

    sheet1 = outputFile.add_sheet("sheet1")

    #Writes the column names
    sheet1.write(0, 0, "Server Name")
    sheet1.write(0, 1, "Instance ID")
    sheet1.write(0, 2, "Instance Type")
    sheet1.write(0, 3, "Associated Volume Names")
    sheet1.write(0, 4, "Volume ID")
    sheet1.write(0, 5, "EBS Volume Size (GB)")
    sheet1.write(0, 6, "Storage Type")
    sheet1.write(0, 7, "State")
    sheet1.write(0, 8, "Business Owner")
    sheet1.write(0, 9, "IOPS")

    index1 = 1

    listOfRunningInstances = []
    setOfBusinessOwners = set()

    for eachInstance in listOfAllInstancesInfo:
        for eachTag in eachInstance['Tags']:
            if eachTag['Key'] == 'Business Owner':
                setOfBusinessOwners.add(eachTag['Value'])
                break

    print(setOfBusinessOwners)

    noBusinessOwnerList = []

    # Fills out the excel sheet if none of the instances have a specified business owner.
    if len(setOfBusinessOwners) == 0:
        for eachInstance in listOfAllInstancesInfo:
            print(eachInstance)
            if eachInstance['State']['Name'] != 'running':
                sheet1.write(index1, 1, eachInstance['InstanceId'])
                sheet1.write(index1, 2, eachInstance['InstanceType'])
                sheet1.write(index1, 7, eachInstance['State']['Name'])
                for eachTag in eachInstance['Tags']:
                    if eachTag['Key'] == 'Name' or eachTag["Key"] == 'name':
                        sheet1.write(index1, 0, eachTag["Value"])
                        break
                ec2 = boto3.session.Session(profile_name=profile).resource('ec2', region_name=region)
                instance = ec2.Instance(eachInstance['InstanceId'])
                volumes = instance.volumes.all()
                indexer = index1
                for eachVol in volumes:
                    sheet1.write(indexer, 5, eachVol.size)
                    sheet1.write(indexer, 6, eachVol.volume_type)
                    sheet1.write(indexer, 9, eachVol.iops)
                    indexer += 1
                for eachDevice in eachInstance['BlockDeviceMappings']:
                    sheet1.write(index1, 3, eachDevice['DeviceName'])
                    sheet1.write(index1, 4, eachDevice['Ebs']['VolumeId'])
                    index1 += 1
                index1 += 1
            else:
                listOfRunningInstances.append(eachInstance)

    else:

        # Fills out the excel sheet for all instances that are not running and have the business owner tag.
        for eachBusinessOwner in setOfBusinessOwners:
            for eachInstance in listOfAllInstancesInfo:
                hasBusinessOwner = False
                for eachTag in eachInstance['Tags']:
                    if eachTag['Key'] == 'Business Owner':
                        hasBusinessOwner = True
                        if eachTag['Value'] == eachBusinessOwner:
                            print(eachInstance)
                            if eachInstance['State']['Name'] != 'running':
                                sheet1.write(index1, 1, eachInstance['InstanceId'])
                                sheet1.write(index1, 2, eachInstance['InstanceType'])
                                sheet1.write(index1, 7, eachInstance['State']['Name'])
                                for eachTag in eachInstance['Tags']:
                                    if eachTag['Key'] == 'Name' or eachTag["Key"] == 'name':
                                        sheet1.write(index1, 0, eachTag["Value"])
                                        break
                                for eachTag in eachInstance['Tags']:
                                    if eachTag['Key'] == 'Business Owner':
                                        sheet1.write(index1, 8, eachTag['Value'])
                                        break
                                ec2 = boto3.session.Session(profile_name=profile).resource('ec2', region_name=region)
                                instance = ec2.Instance(eachInstance['InstanceId'])
                                volumes = instance.volumes.all()
                                indexer = index1
                                for eachVol in volumes:
                                    sheet1.write(indexer, 5, eachVol.size)
                                    sheet1.write(indexer, 6, eachVol.volume_type)
                                    sheet1.write(indexer, 9, eachVol.iops)
                                    indexer += 1
                                for eachDevice in eachInstance['BlockDeviceMappings']:
                                    sheet1.write(index1, 3, eachDevice['DeviceName'])
                                    sheet1.write(index1, 4, eachDevice['Ebs']['VolumeId'])
                                    index1 += 1
                                index1 += 1
                            else:
                                listOfRunningInstances.append(eachInstance)
                        break
                if hasBusinessOwner == False:
                    noBusinessOwnerList.append(eachInstance)

        # Fills out the excel sheet for all instances that are not running and have no business owner tag.
        for eachInstance in noBusinessOwnerList:
            if eachInstance['State']['Name'] != 'running':
                sheet1.write(index1, 1, eachInstance['InstanceId'])
                sheet1.write(index1, 2, eachInstance['InstanceType'])
                sheet1.write(index1, 7, eachInstance['State']['Name'])
                for eachTag in eachInstance['Tags']:
                    if eachTag['Key'] == 'Name' or eachTag["Key"] == 'name':
                        sheet1.write(index1, 0, eachTag["Value"])
                        break
                ec2 = boto3.session.Session(profile_name=profile).resource('ec2', region_name=region)
                instance = ec2.Instance(eachInstance['InstanceId'])
                volumes = instance.volumes.all()
                indexer = index1
                for eachVol in volumes:
                    sheet1.write(indexer, 5, eachVol.size)
                    sheet1.write(indexer, 6, eachVol.volume_type)
                    sheet1.write(indexer, 9, eachVol.iops)
                    indexer += 1
                for eachDevice in eachInstance['BlockDeviceMappings']:
                    sheet1.write(index1, 3, eachDevice['DeviceName'])
                    sheet1.write(index1, 4, eachDevice['Ebs']['VolumeId'])
                    index1 += 1
                index1 += 1
            else:
                listOfRunningInstances.append(eachInstance)

    # Fills out the excel sheet for all instances that are running. The instances are already sorted by business owner.
    for eachInstance in listOfRunningInstances:
        print(eachInstance)
        sheet1.write(index1, 1, eachInstance['InstanceId'])
        sheet1.write(index1, 2, eachInstance['InstanceType'])
        sheet1.write(index1, 7, eachInstance['State']['Name'])
        for eachTag in eachInstance['Tags']:
            if eachTag['Key'] == 'Name' or eachTag["Key"] == 'name':
                sheet1.write(index1, 0, eachTag["Value"])
                break
        for eachTag in eachInstance['Tags']:
            if eachTag['Key'] == 'Business Owner':
                sheet1.write(index1, 8, eachTag['Value'])
                break
        ec2 = boto3.session.Session(profile_name=profile).resource('ec2', region_name=region)
        instance = ec2.Instance(eachInstance['InstanceId'])
        volumes = instance.volumes.all()
        indexer = index1
        for eachVol in volumes:
            sheet1.write(indexer, 5, eachVol.size)
            sheet1.write(indexer, 6, eachVol.volume_type)
            sheet1.write(indexer, 9, eachVol.iops)
            indexer += 1
        for eachDevice in eachInstance['BlockDeviceMappings']:
            sheet1.write(index1, 3, eachDevice['DeviceName'])
            sheet1.write(index1, 4, eachDevice['Ebs']['VolumeId'])
            index1 += 1
        index1 += 1

    outputFile.save(fileName)

    return True

# The function below pulls information about unattached volumes in a specified profile and region.

def volumeInfoPuller(profile, region):

    ec2 = boto3.session.Session(profile_name=profile).resource('ec2', region_name=region)

    volumes = ec2.volumes.all()

    listOfVolumes = []

    for eachVol in volumes:
        listOfVolumes.append(str(eachVol.id))

    listOfUnusedVolumes = []

    for eachVol in listOfVolumes:
        if (ec2.Volume(eachVol).state != 'in-use'):
            listOfUnusedVolumes.append(eachVol)

    return listOfUnusedVolumes

# The function below takes the information from volumeInfoPuller and inputs the information into an excel file.
# A profile and region must be specified. Additionally, a fileName must be specified, and the output excel file is
# saved under that fileName.

def volumeExcelCreator(listOfUnusedVolumes, profile, region, fileName):

    if len(listOfUnusedVolumes) == 0:
        return False

    ec2 = boto3.session.Session(profile_name=profile).resource('ec2', region_name=region)

    outputFile = Workbook()

    sheet1 = outputFile.add_sheet("sheet1")

    #Writes the column names
    sheet1.write(0, 0, "Volume Name")
    sheet1.write(0, 1, "Volume ID")
    sheet1.write(0, 2, "Size(GiB)")
    sheet1.write(0, 3, "Volume Type")
    sheet1.write(0, 4, "IOPS")
    sheet1.write(0, 5, "State")
    sheet1.write(0, 6, "Cost($)")
    sheet1.write(0, 7, "Snapshot Cost/Month")

    index = 1

    # Writes the volume information onto an excel sheet.
    for eachVol in listOfUnusedVolumes:
        print(eachVol)
        volumeInfo = ec2.Volume(eachVol)
        for eachTag in ec2.Volume(eachVol).tags:
            if eachTag['Key'] == 'Name' or eachTag['Key'] == 'name':
                sheet1.write(index, 0, eachTag['Value'])
        sheet1.write(index, 1, eachVol)
        sheet1.write(index, 2, volumeInfo.size)
        sheet1.write(index, 3, volumeInfo.volume_type)
        sheet1.write(index, 4, volumeInfo.iops)
        sheet1.write(index, 5, volumeInfo.state)
        index += 1

    outputFile.save(fileName)

    return True

fileNameList = []

# The block of code below loops through each region and each profile in the regions and profiles list respectively.
# It also differentiates between pulling unattached volume information or server information, depending on the value of
# Server.


for eachProfile in profiles:
    for eachRegion in regions:
        try:
            fileName = eachProfile + "-" + eachRegion + ".xls"
            if Server:
                worked = createExcel(instanceInfoPuller(eachProfile, eachRegion), eachProfile, eachRegion, fileName)
            else:
                worked = volumeExcelCreator(volumeInfoPuller(eachProfile, eachRegion), eachProfile, eachRegion, fileName)
            if worked:
                fileNameList.append(fileName)
            else:
                if Server:
                    print(eachProfile + " + " + eachRegion + " doesn't have any servers.")
                else:
                    print(eachProfile + " + " + eachRegion + " doesn't have any unattached volumes.")
        except Exception as e:
            # Prints if the requested profile and region don't have any servers in them.
            if Server:
                print(eachProfile + " + " + eachRegion + " does not have any servers.")
            else:
                print(eachProfile + " + " + eachRegion + " does not have any unattached volumes.")

# The few lines of code below takes each individual excelFile for each region and each profile and combines them into
# one master excel file.

finalExcelFile = pd.ExcelWriter("OutputFile.xls")

# Takes each individual profile/region file and combines them into one master file.
for eachFile in fileNameList:
    tempFileName = pd.read_excel((eachFile), 'sheet1')
    tempFileName.to_excel(finalExcelFile, sheet_name=eachFile)

finalExcelFile.save()
