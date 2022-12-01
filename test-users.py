# :: this should all be in a .bat file
# :: the program requires Salesforce CLI and Python 3+
# :: 1. authenticate with Salesforce (can be automated with JWT using different command, but I kept it simple for now)
# sfdx force:auth:web:login
# :: 2. get active manual and automation test users in a JSON file
# sfdx force:data:soqlquery -q "SELECT IsActive, Email, Name, Username, Profile.Name, User_Role__c FROM User WHERE IsActive = TRUE" --json
# :: 3. get permission sets in a JSON file
# sfdx force:data:soqlquery -q "SELECT PermissionSet.Name, Assignee.Username FROM PermissionSetAssignment" --json
# :: 4. run python script to transform data
# python3 testUserScript.py
#     :: 1. initialize a Pandas.DataFrame for each JSON file
#     :: 2. testUserDataFrame = join each Pandas.DataFrame using DataFrame.join() function with an INNER JOIN on the username and Assignee.Username fields
#     :: 3. write testUserDataFrame to an excel file with testUserDataFrame.to_excel() function
# :: 5. open the test user excel file
# start testUserData.xlsx

import pandas as pd
import subprocess


def main():
    # get username from terminal argument
    # construct test user excel sheet
    getSpreadsheet("kevin.barreiro@arthrex.com.test")
    # Close the program
    exit()


def getSpreadsheet(username):
    # 1. Login to Salesforce
    subprocess.run("sfdx force:auth:web:login -r \"https://arthrex--test.sandbox.my.salesforce.com/\"", shell=True)

    print("Constructing test user spreadsheet...")
    # 2. Get manual and automation users from Salesforce
    # create a CSV file to store the data
    testUserFile = open("testUser.csv", 'w')
    # Execute the command to retreive the test user data
    subprocess.Popen(f'sfdx force:data:soql:query -u {username} -q "SELECT IsActive, Email, Name, Username, Profile.Name, User_Role__c FROM User WHERE IsActive = TRUE"', shell=True, stdout=testUserFile).wait()
    # close the CSV file
    testUserFile.close()
  
    # 3. Get permission sets from Salesforce
    # create a CSV file to store the data
    permissionSetFile = open("permissionSet.csv", 'w')
    # Execute the command to retreive the test user data
    subprocess.Popen(f'sfdx force:data:soql:query -u {username} -q "SELECT PermissionSet.Name, Assignee.Username FROM PermissionSetAssignment', shell=True, stdout=permissionSetFile).wait() 
    # close the CSV file
    permissionSetFile.close()
    
    # 4. Join the test user data and the permission set data
    # Convert testUserData.csv file into a Pandas.Dataframe
    testUserDataframe = pd.read_fwf("testUser.csv")
    # Convert permissionSetData.csv file into a Pandas.Dataframe
    permissionSetDataframe = pd.read_fwf("permissionSet.csv")
    # Rename the columns in permissionSetDataframe
    columnMapping = {permissionSetDataframe.columns[0]: 'PERMISSIONSET', permissionSetDataframe.columns[1]: 'USERNAME'}
    permissionSetDataframe = permissionSetDataframe.rename(columns=columnMapping)
    # Perform an INNER JOIN on the username and Assignee.Username fields of testUserData_df and permissionSetData_df
    joinedDataframe = testUserDataframe.merge(permissionSetDataframe, how='inner', on='USERNAME')

    # 5. Output the test user excel file
    # Output the joined Pandas.Dataframe to an excel file
    # potentially delete the two csv files at this step
    # Open the excel file


if __name__ == '__main__':
    main()