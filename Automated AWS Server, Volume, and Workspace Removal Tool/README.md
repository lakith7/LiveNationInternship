<img src="https://upload.wikimedia.org/wikipedia/commons/1/1d/AmazonWebservices_Logo.svg" width="400" height="200">

# Amazon Web Services Server, Volume, and Workspace Removal Tool

There are three scripts in this folder

"WorkSpaceInfoCollectorScript.py" collects data on all AWS Workspaces that haven't been accessed in the past 90 days (this value can be changed).

"PullServer:VolumeInfo.py" collects data on all AWS servers that are stopped and all AWS volumes that are unattached.

"EmailAutomator.py" takes the email addresses of the owners of a specified AWS workspace, server, or volume and sends them an email asking whether the AWS workspace, server, or volume can be deleted. The information the script inputs is the same info that is outputed from the other two scripts.

# Inputs

This script takes an input of an excel file full of tagging information for AWS servers. This excel file should simply have column named "Instance ID", which is the ID of the server. Then, there should be multiple other columns with the column name being the tag name, and the column value being the tag value.

The three inputs for this script is as follows: filePath, sheetName, and listOfWantedCategories. 

* filePath is the file path to the excel file. 
* sheetName is the name of the sheet you want to use within the excel file. 
* listOfWantedCategories is a list of all categories from the excel file that you want to have imported in as an AWS server tag. This functionality enables the user to provide a general excel file of servers and tag attributes and only update some of those tags.
  
# Outputs  
   
This script technically has no output. Instead, as it runs, it updates the user as to whether a tag was updated for a specific server or whether it failed. Once it has run, all the server tags should be updated.

## Packages Used

The following packages were used and should be imported before running the specified script:

WorkSpaceInfoCollectorScript.py:

* boto3
* pandas
* datetime
* xlwt
* timedelta
* os
* pytz

PullServer:VolumeInfo.py:

* boto3
* xlwt
* pandas

EmailAutomator.py:

* smtplib
* pandas

## Acknowledgments

* Thank you to Byron Brummer for helping me refine my code and add command line functionality.

## Author

Kamila Wickramarachchi 

[My Github Profile](https://github.com/lakith7)
