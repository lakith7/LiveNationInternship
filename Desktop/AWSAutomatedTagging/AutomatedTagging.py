import pandas as pd
import boto3

def parseInputandCreateTags(filePath, sheetName):

    role = "cs-test"
    region = "us-east-1"

    os.environ['AWS_PROFILE'] = role
    os.environ['AWS_DEFAULT_REGION'] = region
    client = boto3.client('ec2', region_name=region)

    #Reads the file
    file = pd.read_excel(filePath, sheet_name=sheetName)

    listInstanceID = []

    #Adds all instance ID values to a list
    for i in file.index:
        listInstanceID.append(file['Instance ID'][i])

    #Gets name of all columns and adds them to a list
    columns = file.columns
    j = 0

    #Every loop updates all tags for one instance ID. The loop ends when all instance IDs are updated.
    while (j < len(listInstanceID)):
        listOfDicts = []
        listOfResources = []
        i = 0

        #Creates a list that has one item, the Instance ID (Resource name)
        listOfResources.append(listInstanceID[j])

        #Creates a list of individual dictionaries that each have one key value pair for the tags.
        while (i < len(columns)):
            tagsDict = {}
            listOfVals = []
            listOfVals.append(file[columns[i]][j])
            if (type(listOfVals[0]) is float):
                listOfVals[0] = str(listOfVals[0])
            tagsDict["Key: " + columns[i]] = "Value: " + listOfVals[0]
            listOfDicts.append(tagsDict)
            i += 1
        j += 1

        #Creates the tags
        client.create_tags(Resources=listOfResources, Tags=listOfDicts)


#DONT FORGET TO ALTER CODE IN REGARDS TO REGION