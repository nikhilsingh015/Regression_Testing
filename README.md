# Regression_Testing
This is the Selenium script to automate the regression testing using Java


Purpose:
The script was developed to automate the repetitve task of Regression testing.

The process of Regression testing involded the following steps.

Read all the policies from the excel sheet and generate a pdf files for all of them. Then using "Beyond Compare" application compare with the files from the previous release, which is currently in production.

The Entire task is basically to parts:
1) PDF Generation - Implemented using Selenium, Final output is a .jar file and the script will run simply with a double click.
2) PDF Comparision - Implemented using Beyond Compare Script and is executable with the double click on the executable script.

To make it more simpler, I created a Batch Script which will call the forementioned scripts in the sequence.

The whole regression testing is started with a double clck on the batch script and the final results are displayed when the task is completed.
The script runs in parallel and the user can still use the computer or can even lock it and leave for a break.


Functionalities covered in this script:

1) Logging.
2) Download PDF Functionality.
3) Creating Random Folder Name for different Products.
4) Read the Policy number from Excel Sheet.
5) Login to the website using UserName and Password.
6) Switch tabs and close unnecessary tabs.
7) Loop to generate policies for different products as listed in the excel sheet.
8) Creating random proposal name.
9) Open the created policy in new tab and save it as pdf with a standard identifiable name.
10) Check if the pdf file is downloaded successfully and print the success message in the logs.
11) If policy download fails, save it in a list and reattempt it later.
12) Delete the generated policies from the website to remove the proposals generated in each iteration.
13) Reattempt for the failed cases and then print the final status, and print the policy number if failed for any in the last attempt.
14) Move the files into a new directory which will then be used for pdf comparision in the "Beyond Compare" Application.

