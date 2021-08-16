# report-automation
Automated report generation with a click of a button in your sheet

Instructions:
First, open the 'source data.xlsm' file. Go to Developer Tab -> Visual Basic. You will see the VBA window open.
Go to Modules/module1 file and change the variable for python path and script path for your PC.

Python path can be found using `$where python` command in cmd.

Now, change the in_filename and out_filename variables in Python script in lines 6, 7 to the source data.xlsm and report template.xlsx files.
Next, save VBA file and click on the Generate Report Button in the source data XLSM file.
Your report will be generated with the required data in the same folder as of script.
