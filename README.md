US GPS Deployment and MySource are two internal Deloitte web sites that contain listings of open roles within Deloitte.  This repository contains several Python scripts designed to make the job search more efficient for the searcher.

The script process requires the searcher to download a spreadsheet from either or both of the web sites and save it on the searrcher's local manchine.

For US GPS Deployment, see https://resources.deloitte.com/sites/growth/Industries/USGPS/Pages/GPS_resource_management. Click on "Read More" in the Roles/Demands box, and then the "GPS open positions" link to download the spreadsheet.  

For MySource, see https://mysource1.deloitte.com/Requests/Search.

Note for MySource download: Since the spreadheet downloaded from MySource does not contain some of the needed information, it is advantageous to make some of the selections under "Advanced Search" - on the web page.  For US-based practitioners, under "Practice", select all of the listed options - except any of the "USI" practices.  (This selection by itself eliminates close to 85% of the roles listed.)  Unless you have an active Security Clearance, select "NA" under "Security Clearance Required".  Note: The GPS Open Demands worksheet contains "Request Paractice" and "Security Clearance Required" columns, so these selections can be done as part of the script processing, and need not be done on the web page.

The following are short descriptions of the modules:

MyLocation.py - Uses an input spreadsheet tailored to the searcher, to determine the proximity of the project location - for use with the Co-Location value.

MySource.py - Uses configuration information and a dowmloaded spreadsheet, to select only the roles that match the criteria selected by the searcher.  Roles are identified by the appropriate filter category, and roles are matched to previous action (applied, passed over, or other).

Apply.py - Displays some information about a specific role - identified by Request Id - and facilitates updates to the Action file for the seatched.  (The Action file is used, in later cycles, by MySource.py, to display previous action for matched roles.)

Mario.py - Utility program to update the searcher's Action file - if the searcher had not been using Apply.py or otherwise updated the Action file.

ReadControl.py - Utility script for reading a control file - used by MyLocation.py, MySource.py, Apply.py, and Mario.py.

FilterRow.py - Script for handling a spreadsheet of filter criteria, to by used by MySource.py.  (Work still in progress, prior to initial implementation.)
