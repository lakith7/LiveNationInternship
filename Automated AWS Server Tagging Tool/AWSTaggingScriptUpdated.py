import pandas as pd
import boto3
import os
import xlrd
import sys
import time
import argparse
import jmespath
import botocore


# READ BEFORE RUNNING:
# All arguments are case sensitive and white space sensitive. Make sure they are all correct before running.
# Filepath is the full filepath to the excel file containing the AWS tagging information.
# Sheetname is the sheetname in the excel file that you want to use. It is found at the bottom of the file.
# listOfWantedCategories is a list of all the categories that you want to be tagged. It must be in the form
# ["firstCategory", "secondCategory", "thirdCategory"] . Feel free to put as many categories you want. Don't add "AZ" or
# "Account Name" to the list as those are pulled even if they are not in the listOfWantedCategories. Finally, make sure
# each row in the sheet you are using has the exact same "Account Name". Otherwise, the program will display an error
# message.

def main():
    args = get_args()
    parseInputandCreateTags(args.file, args.sheet, args.categories)


def get_args():
    parser = argparse.ArgumentParser()
    parser.add_argument('-f', '--file',
                        type=str,
                        default='/Users/kamila.wickramarachchi/Desktop/Scripts/AWSTaggingPatchWindow/CSInventoryPatchWindow.xlsm',
                        help='Excel file containing the AWS tagging information'
                        )
    parser.add_argument('-s', '--sheet',
                        type=str,
                        default='CS Prod',
                        help='Sheetname in the excel file that you want to use'
                        )
    parser.add_argument('-c', '--categories',
                        type=str,
                        nargs='*',
                        default=['Patch Window'],
                        help='List of all the categories that you want to be tagged'
                        )

    args = parser.parse_args()
    return args

def parseInputandCreateTags(filePath, sheetName, listOfWantedCategories):
    # Reads the file
    file = pd.read_excel(filePath, sheet_name=sheetName)

    listInstanceID = []

    # Adds all instance ID values to a list
    for i in file.index:
        listInstanceID.append(file['Instance ID'][i])

    # Gets name of all columns and adds them to a list
    columns = file.columns

    j = 0

    # Every loop updates all tags for one instance ID. The loop ends when all instance IDs are updated.

    for id in listInstanceID:

        listOfDicts = []

        instanceID = id

        # Creates a list of individual dictionaries that each have one key value pair for the tags.
        for col in columns:
            tagsDict = {}
            listOfVals = []
            listOfVals.append(file[col][j])

            listOfVals[0] = str(listOfVals[0])
            if col and listOfVals[0] and (col != "nan") and (listOfVals[0] != "nan"):
                if (str(col) in listOfWantedCategories):
                    tagsDict["Key"] = str(col)
                    tagsDict["Value"] = str(listOfVals[0])
                    listOfDicts.append(tagsDict)
        j += 1

        tagMethod(instanceID, listOfDicts)

# Takes input information and creates a tag.
def tagMethod(instanceID, listOfDicts):
    # Review - AWS_DEFAULT_REGION; You're calling out the region_name in the client() setup, so
    # there's no need (or use) to setting the env variable as it's overruled.
    # Review - AWS_PROFILE; Similarly you can call out the profile here in code like this:
    #
    #          client = boto3.session.Session(profile_name = role).client('ec2', region_name=region)
    #
    #          Although I'd *strongly* recommend not switching profiles like
    #          this at all in code as a best practice.

    # Set profile before running the code. Also pull all regions and run through all of them, using a try and catch
    # to grab the one's that error.
    region = "us-east-1"

    client = boto3.client('ec2', region_name=region)

    # If I try using client = boto3.client('ec2', region_name=region), there is an error stating I don't have the
    # authorization to run describe_region(). The weird thing is that I can run describe_region() in the command line
    # with no problems.
    #
    # Also I start with region = 'us-east-1' because I needed to specify a region in order to run
    # describe_regions().
    #
    # To test I was using
    # client = boto3.session.Session(profile_name = role).client('ec2', region_name=region) because it was the only way
    # to get my code to run.

    # Retrieves all regions/endpoints that work with EC2
    response = client.describe_regions()
    regionInfo = response['Regions']
    for eachRegion in regionInfo:
        try:
            region = eachRegion['RegionName']
            client = boto3.client('ec2', region_name=region)
            client.create_tags(Resources=[instanceID], Tags=listOfDicts)
        except botocore.exceptions.ClientError:
            print("Trying another region")


# Below method is not used. I just want to keep the code in case it is useful in the future.
def addTagsbyRegion(filePath, sheetName):
    file = pd.read_excel(filePath, sheet_name=sheetName)
    regionSet = set()
    roleSet = set()

    for i in file.index:
        print(str(file["AZ"][i]))
        print(str(file["Account Name"][i]))
        regionSet.update([str(file['AZ'][i])])
        roleSet.update([str(file["Account Name"][i])])

    for eachRole in roleSet:
        for eachRegion in regionSet:
            parseInputandCreateTags(filePath, sheetName)


if __name__ == "__main__":
    main()
