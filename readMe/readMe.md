## TL;DR
### **Before Running the Script:**
  - Fill out Config File
  - Add users and message variables to CSV Template
  - Add the desired External Auto Reply to externalMessage.txt
  - Don't change dependency file names
  - Keep Dependency files in the same folder as the script

## Description
Set Users mailbox's external auto reply to the content added to the externalMessages.txt file.

## Function of the Script
The script can set Auto Reply including the External Message for bulk users. The script will also get the current configuration of each user and save it to a log file before making any changes. All actions the script takes are logged.

## Dependencies
- ExchangeOnlineManagement Module
  - There is a check at the beginning of the script that will import/install the module if it isn't detected.
- Config.json
  - Used for declaring EXO Admin Credentials
- userImport.CSV
  - Place user mailbox address in the 'user' column. Add additional columns for variables to be used in externalMessage.text.
  - **WARNING:** Do not rename 'user' column

- ExternalMessage.txt
    - Create external out-of-office message in this file. Variables can be used as long as they are wrapped by a ``` $()```. When adding new variables, the variable name should be a column header and the value in the rows below the header.
  - examples:
    ![userImport.csv ScreenCap](assets\userImportExample.png)

  **Syntax:** ```This is a test message for $($_.UserName).```
  
  **Result:** This is a message for Jim Bob.

## Instructions
1. Keep all script dependency files in the same folder as the script
2. Don't change the name of the dependency files
3. Complete the Configuration file
4. Add users to the userImport.csv file and any variables needed in the External Message
5. Add message to the externalMessage.txt file.
6. Run as Administrator if its possible the ExchangeOnlineManagement module isn't installed on host