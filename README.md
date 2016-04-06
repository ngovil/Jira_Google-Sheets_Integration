# Jira_Google-Sheets_Integration

This is a Google add-on that is able to automate the transfer and updating of the JIRA issues into a Google Sheet.

It has been created for the Naehas and is set uniquely for that company, however it is possible to change the code to adjust for different companies. You would need to change the host name (it is currently set at "naehas.jira.com:), the project that contains the keys you wish to import (it is currently "FRED"), you would need to remove the sections that corrospond to any custom fields, and add any custom fields that your project uses. Please feel free to email me at neha.j.govil@gmail.com for any further help.

The first time this add-on is used on a particular sheet, it must be configured by inputting the categories to be imported and their corresponding heading names. After entering the credentials, all that must be done is press “Refresh Now” and the data in the sheet will refresh. The data is drawn through using URL fetch to collect the information from the JIRA Rest API. After that, there is a lot of parsing and rearranging in order to display the data appropriately, based on your configurations. 

How to install this script on a sheet:

  You first access the Google Sheet where you wish to add the script.
  
  Go to Tools => Script Editor...

  Create a script for Blank Project.

  Delete all the code currently in the page and paste in the code from the code.gs page.

  Save the project.
  
  Go back to the Google Sheet and reload.

  There should now be another category in the navigation bar called "JIRA".

  You are now ready to begin using the add-on.



