# borsheimsSite-test
I've utilized Selenium Java to test the "Borsheims" site using Eclipse IDE.

## **Test Case 1**
- Open the site and log in with a valid account
- Search a list of brands (Hobo, Citizen, Baccarat).
- *This list should be stored in an Excel file and the code chooses the brand randomly.*
- *Make sure the brand name is displayed as a title on the search result page.*
- Add the item to the bag and navigate to the bag page.
- Display the list of the items on the bag page.

## **Test Case 2**
- Open the site and go to the Create Account page.
- Do a registration for new users from an Excel file that you will fill in.
- try creating 3 new users, and once you create the user, add "X" to the Excel sheet so the user information will not be used again.
- Each user should be created in a separate window and alone.
- Make sure the registration is completed successfully and the username is displayed.
- Log out from the account and log in again.

📎 Note: The test cases mentioned above should exist in the same Excel file for user information, but in another tab (separated tabs) for the test case you will execute, please mark it as "X" and let the code run the test case which has the mark "X”

📎 In the Excel file, you should have a "Status" column and this should be filled from the code based on the status that you will have from the test run.
the values should be "Passed or Failed"
