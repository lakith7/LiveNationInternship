import pandas as pd
from xlwt import Workbook
import xlrd

# IMPORTANT! PLEASE READ BEFORE RUNNING:
#
# For both the svrpatch servers and the cloudservices servers, once you pull the info as a csv from the active
# directory, delete the first row. That first row contains info not necessary for this script and in fact will break the
# script. Make sure that the first row is the column names and not some other information.
#
# The first argument is the filepath to the tanium servers csv, the second argument is the filepath to the svrpatch servers
# csv, and the third argument is the filepath to the cloud services OU servers csv.
#
# The next three arguments each require a string that is the name of the column under which the server names are
# kept in the respective csv file. This argument is case sensitive and whitespace sensitive, so make sure to put
# the EXACT same column name.
#
# Fill in the variables below. I left in prefilled values just to show the format for each argument.
# Make sure to put your values in place of mine.
#
# The output excel file is saved under the name "ServerInfo.xls" and is saved wherever this python script is saved/run.

TaniumServerFilePath = "/Users/kamila.wickramarachchi/Desktop/DamionServerInformation/ListOfAllTaniumServers.csv"
SvrPatchServersFilePath = "/Users/kamila.wickramarachchi/Desktop/DamionServerInformation/SvrPatchServers.csv"
CloudServicesServersFilePath = "/Users/kamila.wickramarachchi/Desktop/DamionServerInformation/CloudServicesServers.csv"
TaniumServerColumnName = "Computer Name"
SvrPatchServersColumnName = "name"
CloudServicesServersColumnName = "Name"

# The below function takes in server information from tanium, the svrpatch security group, and cloud services and
# cross checks the servers and then inputs them into an excel file.
def svrPatchCrossCheck(filePathofTaniumServers, filePathofSvrPatch, filePathofCloudServices, taniumColumnName,
                       svrPatchColumnName, cloudServicesColumnName):

    # The below code reads each csv file and stores the information.
    taniumServers = pd.read_csv(filePathofTaniumServers)
    svrPatchServers = pd.read_csv(filePathofSvrPatch)
    cloudServicesServers = pd.read_csv(filePathofCloudServices)

    outputFile = Workbook()

    sheet1 = outputFile.add_sheet("sheet1")

    listOfTaniumServers = []
    listOfCloudServicesServers = []
    listOfSvrPatchServers = []

    # The below code creates lists of all servers in Tanium, the svrPatch security group, and Cloud Services.

    for eachServer in taniumServers.index:
        listOfTaniumServers.append(taniumServers[taniumColumnName][eachServer])

    for eachServer in svrPatchServers.index:
        listOfSvrPatchServers.append(svrPatchServers[svrPatchColumnName][eachServer])

    index1 = 1

    for eachServer in cloudServicesServers.index:
        listOfCloudServicesServers.append(cloudServicesServers[cloudServicesColumnName][eachServer])
        sheet1.write(index1, 4, str(cloudServicesServers["CanonicalName"][eachServer]))
        sheet1.write(index1, 5, str(cloudServicesServers["Enabled"][eachServer]))
        sheet1.write(index1, 6, str(cloudServicesServers["MemberOf"][eachServer]))
        sheet1.write(index1, 7, str(cloudServicesServers["PasswordLastSet"][eachServer]))
        index1 += 1

    index = 0

    # Strips off the .LYV.LiveNation.com part of the server name.
    for eachServer in listOfTaniumServers:
        if ".lyv.livenation.com" in str(eachServer) or ".LYV.LiveNation.com" in str(eachServer):
            listOfTaniumServers[index] = eachServer[:-19]
        index += 1

    # Writes the column names.
    sheet1.write(0, 0, "Server Name")
    sheet1.write(0, 1, "In SvrPatch Patch Group?")
    sheet1.write(0, 2, "In Tanium?")
    sheet1.write(0, 3, "In CloudServicesOU?")
    sheet1.write(0, 4, "CanonicalName")
    sheet1.write(0, 5, "Enabled")
    sheet1.write(0, 6, "MemberOf")
    sheet1.write(0, 7, "PasswordLastSet")

    index = 1

    # The below code takes all server information, cross checks them, and then saves the information to an excel sheet.
    for eachServer in listOfCloudServicesServers:
        sheet1.write(index, 0, str(eachServer))
        if eachServer in listOfCloudServicesServers:
            sheet1.write(index, 3, "Yes")
        else:
            sheet1.write(index, 3, "No")
        if eachServer in listOfSvrPatchServers:
            sheet1.write(index, 1, "Yes")
        else:
            sheet1.write(index, 1, "No")
        if eachServer in listOfTaniumServers:
            sheet1.write(index, 2, "Yes")
        else:
            sheet1.write(index, 2, "No")
        index += 1

    for eachServer in listOfSvrPatchServers:
        if eachServer not in listOfCloudServicesServers:
            sheet1.write(index, 0, str(eachServer))
            if eachServer in listOfCloudServicesServers:
                sheet1.write(index, 3, "Yes")
            else:
                sheet1.write(index, 3, "No")
            if eachServer in listOfSvrPatchServers:
                sheet1.write(index, 1, "Yes")
            else:
                sheet1.write(index, 1, "No")
            if eachServer in listOfTaniumServers:
                sheet1.write(index, 2, "Yes")
            else:
                sheet1.write(index, 2, "No")
            index += 1

    for eachServer in listOfTaniumServers:
        if (eachServer not in listOfCloudServicesServers) & (eachServer not in listOfSvrPatchServers):
            sheet1.write(index, 0, str(eachServer))
            if eachServer in listOfCloudServicesServers:
                sheet1.write(index, 3, "Yes")
            else:
                sheet1.write(index, 3, "No")
            if eachServer in listOfSvrPatchServers:
                sheet1.write(index, 1, "Yes")
            else:
                sheet1.write(index, 1, "No")
            if eachServer in listOfTaniumServers:
                sheet1.write(index, 2, "Yes")
            else:
                sheet1.write(index, 2, "No")
            index += 1

    outputFile.save("ServerInfo.xls")
    print("Finished!")

svrPatchCrossCheck(TaniumServerFilePath, SvrPatchServersFilePath, CloudServicesServersFilePath, TaniumServerColumnName,
                   SvrPatchServersColumnName, CloudServicesServersColumnName)

