<p align="middle">
  <img src="https://upload.wikimedia.org/wikipedia/commons/1/1d/AmazonWebservices_Logo.svg" width="250" />
  <img src="https://jumpcloud.com/wp-content/uploads/2017/12/jc-sso-datadog.png" width="250" /> 
  <img src="https://www.tanium.com/uploads/Tanium-Logo-FullColor-Positive.jpg" width="250" />
</p>

# Amazon Web Services Server, Volume, and Workspace Removal Tool

There are two scripts in this folder.

The active directory is a place where all server information was saved for the company.

"DDAWSComparer.py" compares the servers in the company's active directory, DataDog, and Amazon Web Services and outputs an excel file that states for each server whether it is in the active directory, DataDog, AWS, or a combination of the three.

"SvrPatchCrossChecking.py" compares the servers in Tanium, the active directory under Cloud Services, and the active directory under SvrPatch and outputs an excel file that states for each server whether it is in Tanium, the active directory under Cloud Services, the active directory under SvrPatch, or a combination of the three.

# Inputs

DDAWSComparer.py:

* APIKey: A string that is required for DataDog API access. This key varies for each company. 
* APPKey: A string that is required for DataDog API access. This key varies for each company. 
* filePath: A string that contains the file path to wherever you want the output excel file to be saved.
* CSInventoryFilePath: A string that contains the file path to a list of servers that was given to me to compare all data against. This excel file cannot be recreated (by me at least), as it was given to me by my boss.
* ADServerInfoFilePath: A string that contains the file path to wherever the server information from the active directory is stored. The one line script to access this information is in the README sections of the DDAWSComparer.py script.

SvrPatchCrossChecking.py:

* TaniumServerFilePath = A string that contains the filePath to an excel file that contains information about servers in Tanium. This file was pulled from the Tanium online client.
* SvrPatchServersFilePath = A string that contains the filePath to an excel file that contains information about servers in the SvrPatch category of the active directory. This file was pulled from the company's active directory.
* CloudServicesServersFilePath = A string that contains the filePath to an excel file that contains information about servers in the Cloud Services category of the active directory. This file was pulled from the company's active directory.
* TaniumServerColumnName = A string that contains the column name in the Tanium servers file that contains the server ID.
* SvrPatchServersColumnName = A string that contains the column name in the SvrPatch servers file that contains the server ID.
* CloudServicesServersColumnName: A string that contains the column name in the Cloud Services servers file that contains the server ID.

# Outputs  
   
DDAWSComparer.py:

* Outputs a single excel file that states for each server whether it is in the active directory, DataDog, AWS, or a combination of the three.

SvrPatchCrossChecking.py:

* Outputs an excel file that states for each server whether it is in Tanium, the active directory under Cloud Services, the active directory under SvrPatch, or a combination of the three.

## Packages Used

The following packages were used and should be imported before running the specified script:

DDAWSComparer.py:

* boto3
* os
* xlwt
* pandas
* datadog
* requests
* json
* csv

SvrPatchCrossChecking.py:

* pandas
* xlwt
* xlrd


## Acknowledgments

* Thank you to Byron Brummer for helping me refine my code.
* Thank you to Damion Sandidge for helping me navigate the active directory and learn Powershell.

## Author

Kamila Wickramarachchi 

[My Github Profile](https://github.com/lakith7)
