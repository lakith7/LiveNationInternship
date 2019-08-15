import smtplib
from smtplib import SMTPException
import pandas as pd

# READ BEFORE RUNNING:
#
# There is two input values for this script.
#
# The first input value (workspaceScriptFilePath) is the file path to the xls output of the
# WorkSpaceInfoCollectorScript.py.
#
# The second input value is adOutputFilePath, which is simply the file path to the output from the active directory
# script.
# Make SURE to delete the first row of the output file, as it is unnecessary data that will break the script.
# The active directory csv file can be retrieved by running the following script for each profile and region
# in the active directory (powershell):
#
#  $List = @('Rob.Caudillo', 'vivian.wang', 'Venu.Malyala', 'Christine.Chu', 'Matt.Jensen', 'Stacy.Price', 'curtis.li', 'Kushal.Gupta', 'kaushlendra.bais', 'jquiring')
#  $Users = @()
#  $Incorrect = @()
#  foreach ($Item in $List) {
#      try {$Users += Get-ADUser $Item -Properties mail,manager,enabled} catch{$Incorrect += $Item}
#  }
#  $Users | Export-Csv -Path .\Desktop\UserInfo.csv
#  Out-File -FilePath .\Desktop\Incorrect.txt -InputObject $Incorrect
#
# Replace the elements of $List with the list of users that is outputted in the
# "List of Usernames (For Active Directory Script)" column in the output excel file of the
# WorkSpaceInfoCollectorScript.py

# FilePath for the output of the WorkSpaceInfoCollectorScript.py
workspaceScriptFilePath = "/Users/kamila.wickramarachchi/Desktop/FinalResult.xls"

# FilePath for the output of the active directory.
adOutputFilePath = "/Users/kamila.wickramarachchi/Desktop/UserInfo.csv"

workspaceScriptInformation = pd.read_excel(workspaceScriptFilePath, sheet_name = "")

adOutputInformation = pd.read_csv(adOutputFilePath)

listOfUsernames = []
listOfEmail = []
listOfFirstName = []
listOfFullName = []
listOfDates = []

#Compiles a list for each input value needed.
for eachUser in adOutputInformation.index:
    listOfUsernames.append(str((adOutputInformation["SamAccountName"][eachUser])))
    listOfEmail.append(str(adOutputInformation["mail"][eachUser]))
    listOfFirstName.append(str(adOutputInformation["GivenName"][eachUser]))
    listOfFullName.append(str(adOutputInformation["GivenName"][eachUser]) + " " + str(adOutputInformation["Surname"][eachUser]))
    for eachWorkspaceUser in workspaceScriptInformation.index:
        if str((workspaceScriptInformation["Username"][eachWorkspaceUser])) == str((adOutputInformation["SamAccountName"][eachUser])):
            listOfDates.append(str(workspaceScriptInformation["User Last Active"][eachUser])[:-22])
            break

index = 0

# The below code creates and then sends an email to every email in the email list. The if clause contains a message that
# will be sent if the last accessed date is null, and the else clause contains a message that will be sent if the last
# accessed date is an actual date.
while (index < len(listOfUsernames)):
    if (listOfDates[index]) == '':
        sender = 'kamilawickramarachchi@livenation.com'
        receivers = [listOfEmail[index]]

        message = """From: Cloud Services Operations <cloudservicesoperations@livenation.com>
To: {0} <{1}>
Subject: AWS Workspace

Hello {2},

I’m reaching out to you about your AWS workspace {3}. 
Since a monthly charge is incurred, we would like to know if this workspace is still being utilized.

Please complete this quick Microsoft form to retain or delete your AWS Workspace: https://tinyurl.com/yxggaj6m

If you have any questions, please send an email to cloudserviceshelp@livenation.com. Please do not reply to this email.

Thank you,

Live Nation Cloud Services
    """
        message = message.format(listOfFullName[index], listOfEmail[index], listOfFirstName[index], listOfUsernames[index])
        message = message.replace(u"\u2019", "'")

        try:
            smtpObj = smtplib.SMTP('mail-aws.lyv.livenation.com')
            smtpObj.sendmail(sender, receivers, message)
            print("Successfully sent email")
        except SMTPException:
            print("Error: unable to send email")

        index += 1
    else:
        sender = 'kamilawickramarachchi@livenation.com'
        receivers = [listOfEmail[index]]

        message = """From: Cloud Services Operations <cloudservicesoperations@livenation.com>
To: {0} <{1}>
Subject: AWS Workspace

Hello {2},

According to our records the last time you accessed your AWS Workspace {3} was on {4}. 
Since a monthly charge is incurred, we would like to know if this workspace is still being utilized.

Please complete this quick Microsoft form to retain or delete your AWS Workspace: https://tinyurl.com/yxggaj6m

If you have any questions, please send an email to cloudserviceshelp@livenation.com. Please do not reply to this email.

Thank you,

Live Nation Cloud Services
        """
        message = message.format(listOfFullName[index], listOfEmail[index], listOfFirstName[index],
                                 listOfUsernames[index], listOfDates[index])
        message = message.replace(u"\u2019", "'")

        try:
            smtpObj = smtplib.SMTP('mail-aws.lyv.livenation.com')
            smtpObj.sendmail(sender, receivers, message)
            print("Successfully sent email")
        except SMTPException:
            print("Error: unable to send email")

        index += 1