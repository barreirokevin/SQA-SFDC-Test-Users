import pandas as pd
import subprocess
import pip
import sys


def main():
    # get the username argument from terminal
    username = sys.argv[1]
    # install dependecies
    pip.main(["install", "openpyxl", "pandas"])
    # construct the test user excel sheet
    getSpreadsheet(username)
    # Close the program
    exit()


def getSpreadsheet(username):
    # 1. Login to Salesforce
    subprocess.run("sfdx force:auth:web:login -r \"https://arthrex--test.sandbox.my.salesforce.com/\"", shell=True)

    # 2. Get manual and automation users from Salesforce
    # create a tmp directory
    subprocess.run("mkdir tmp", shell=True)
    # create a CSV file to store the data
    testUserFile = open("./tmp/testUser.csv", 'w')
    # Execute the command to retreive the test user data
    subprocess.Popen(f'sfdx force:data:soql:query -u {username} -q "SELECT IsActive, Email, Name, Username, Profile.Name, User_Role__c FROM User WHERE IsActive = TRUE"', shell=True, stdout=testUserFile).wait()
    # close the CSV file
    testUserFile.close()
  
    # 3. Get permission sets from Salesforce
    # create a CSV file to store the data
    permissionSetFile = open("./tmp/permissionSet.csv", 'w')
    # Execute the command to retreive the test user data
    subprocess.Popen(f'sfdx force:data:soql:query -u {username} -q "SELECT PermissionSet.Name, Assignee.Username FROM PermissionSetAssignment', shell=True, stdout=permissionSetFile).wait() 
    permissionSetFile.close()
    
    # 4. Join the test user data and the permission set data
    # Convert testUserData.csv file into a Pandas.Dataframe
    testUserDataframe = pd.read_fwf("./tmp/testUser.csv")
    # Convert permissionSetData.csv file into a Pandas.Dataframe
    permissionSetDataframe = pd.read_fwf("./tmp/permissionSet.csv")
    # Rename the columns in permissionSetDataframe
    columnMapping = {permissionSetDataframe.columns[0]: 'PERMISSIONSET', permissionSetDataframe.columns[1]: 'USERNAME'}
    permissionSetDataframe = permissionSetDataframe.rename(columns=columnMapping)
    # Perform an INNER JOIN on the username and Assignee.Username fields of testUserData_df and permissionSetData_df
    joinedDataframe = testUserDataframe.merge(permissionSetDataframe, how='inner', on='USERNAME')

    # 5. Output the test user excel file
    # Output the joined Pandas.Dataframe to an excel file
    print("Constructing test user spreadsheet...")
    with pd.ExcelWriter("testUsers.xlsx", engine="openpyxl") as writer:
        joinedDataframe.to_excel(writer)
    # Delete the tmp folder
    subprocess.run("rmdir tmp /s /q ", shell=True)
    # Open the excel file
    subprocess.run("start testUsers.xlsx", shell=True)


if __name__ == '__main__':
    main()