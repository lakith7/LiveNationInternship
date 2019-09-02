<img src="https://assets.pcmag.com/media/images/514204-amazon-web-services-logo.jpg?width=333&height=200" width="333" height="245">

# Amazon Web Services Server Tagging Tool

This script takes an excel file that contains AWS Server tagging information and updates server tags in the AWS console. This script is meant to be run from the command line, and functionality has been added to make command line access as easy as possible. As of 9/2/19, this script has been run on thousands of servers at one time and has fully run within a minute.

# Inputs

This script takes an input of an excel file full of tagging information for AWS servers. This excel file should simply have column named "Instance ID", which is the ID of the server. Then, there should be multiple other columns with the column name being the tag name, and the column value being the tag value.

The three inputs for this script is as follows: filePath, sheetName, and listOfWantedCategories. 

* filePath is the file path to the excel file. 
* sheetName is the name of the sheet you want to use within the excel file. 
* listOfWantedCategories is a list of all categories from the excel file that you want to have imported in as an AWS server tag. This functionality enables the user to provide a general excel file of servers and tag attributes and only update some of those tags.
  
# Outputs  
   
This script technically has no output. Instead, as it runs, it updates the user as to whether a tag was updated for a specific server or whether it failed. Once it has run, all the server tags should be updated.

## Packages Used

The following packages were used and should be imported before running the script:

* pandas
* boto3
* os
* xlrd
* sys
* time
* argparse
* jmespath
* botocore

## Acknowledgments

* Thank you to Byron Brummer for helping me refine my code and add command line functionality.

# Author

Kamila Wickramarachchi 

[My Github Profile](https://github.com/lakith7)
