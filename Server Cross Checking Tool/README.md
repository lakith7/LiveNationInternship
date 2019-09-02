            | 
:-------------------------:|:-------------------------:
![](https://upload.wikimedia.org/wikipedia/commons/1/1d/AmazonWebservices_Logo.svg)  |  ![](https://imgix.datadoghq.com/img/about/presskit/logo-h/logo_horizontal_white.png)  |  ![](https://www.tanium.com/uploads/Tanium-Logo-FullColor-Positive.jpg)

<img src="https://upload.wikimedia.org/wikipedia/commons/1/1d/AmazonWebservices_Logo.svg" width="400" height="200">
<img src="https://imgix.datadoghq.com/img/about/presskit/logo-h/logo_horizontal_white.png" width="400" height="200">
<img src="https://www.tanium.com/uploads/Tanium-Logo-FullColor-Positive.jpg" width="400" height="200">


# Amazon Web Services Server, Volume, and Workspace Removal Tool

There are three scripts in this folder.

"WorkSpaceInfoCollectorScript.py" collects data on all AWS Workspaces that haven't been accessed in the past 90 days (this value can be changed).

"PullServer:VolumeInfo.py" collects data on all AWS servers that are stopped and all AWS volumes that are unattached.

"EmailAutomator.py" takes the email addresses of the owners of a specified AWS workspace, server, or volume and sends them an email asking whether the AWS workspace, server, or volume can be deleted. The information the script inputs is the same info that is outputed from the other two scripts.

# Inputs

WorkSpaceInfoCollectorScript.py:

* role: This is simply a string that is the exact name of the AWS profile you want this script to run in.
* region: This is a string that is the exact name of the AWS region you want this script to run in.
* time: This is an integer that is used to determine which workspaces should be retrieved. If the value is 90 (which the default), then the script pulls all workspaces that haven't been accessed in the last 90 days.

PullServer:VolumeInfo.py:

* Server: A boolean input. If True, the script looks for stopped AWS servers and outputs their metadata. If False, the script looks for unattached AWS volumes and outputs their metadata.
* profiles: A list of strings that contains all the AWS profiles that the user would like to loop through.
* regions: A list of strings that contains all the AWS regions that the user would like to loop through.

EmailAutomator.py:

* workspaceScriptFilePath: A string that is the file path of the output of the WorkSpaceInfoCollectorScript.py
* adOutputFilePath: A string that is the file path for the output from the active directory script (was not allowed to post this script on my github). A sample active directory script is placed in the comments at the top of EmailAutomator.py
* msg: A string that contains the email body message.

Note: EmailAutomator.py was optimized for the WorkSpaceInfoCollector.py script but was able to be used for the PullServer:VolumeInfo.py script with a few minor changes.
  
# Outputs  
   
WorkSpaceInfoCollectorScript.py:

* Outputs an excel file that contains information about all unused AWS workspaces.
* Outputs a list of strings that are each a username. This output was then used in an active directory script to pull email addresses. I was not allowed to publish the active directory script on my github.

PullServer:VolumeInfo.py:

* Outputs an excel file that contains information about all the unattached volumes or stopped servers.

EmailAutomator.py:

* Technically this script has no output. However, as the script attempts to send each email, the user is notified whether the email was succesfully sent or not.

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

* Thank you to Byron Brummer for helping me refine my code.
* Thank you to Tim Cordell for introducing me to the SMTP protocols.

## Author

Kamila Wickramarachchi 

[My Github Profile](https://github.com/lakith7)
