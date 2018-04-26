4/10/2017
Automated Appointment Reminder System
Written by Michael Williams (mrwilliams@seton.org)

Projects and Submodules:

	AppointmentReminder2.4 Update(6/5/2017)

		*Centricity report identification was changed substantially. In the previous version, the two most recently 
		edited (created) files in the Centricity folder were compared by size and the larger sized file was selected 
		to pull patient data. This version loops through all files in the Centricity folder and pulls patient data 
		from any number of files created on the same date the program is run. If multiple files contain patient data, 
		the data from these files can all be uploaded into REDCap.

		*Patient name and location are now appropriately cased instead of using only all caps.

		*Location book data extraction has been modified to allow use of a single standard workbook for both this 
		application and project team daily documentation needs.

		*Added filePath.py file to consolidate all file paths used in the program.

		*Removed dates for sanity checks from references_Text.py and placed in the apptRemind_Text.py file to prevent 
		confusion about where these values are created.

		*Both output log and system log were moved to the shared drive for remote monitoring

		*Appttext.xml file was exported from Task Scheduler to assist with setup when the host machine is refreshed.

	AppointmentReminder2.3 is the main repository for this application (4/10/2017).

		*Log is a file that contains an excel file which lists all of the patients who have received
		an appointment reminder and when the system was run.

		*redcap is a file which contains all of the code needed to implement the redcap API. This was downloaded
		directly from Github in its current form and should not be edited.

		*api_keys.txt just contains the api key needed to submit the records to the appropriate Redcap project.

		*apptRemind_Text.py is the main file from which everything is run. When triggered, it pulls data from
		Centricity reports and appointment location excel files located on the G drive, cobbles the data into
		a json format, logs the data in the log file then submits the data to redcap via the API. 

		*emaillog.py is a submodule that contains the code necessary to open up Microsoft Outlook and send emails.
		This function is activated whenever an error occurs during a program run and is the main logging function
		that enables timely trouble shooting.

		*Program_Start.bat is a batchfile that contains two arguments necessary to run the program (1. the appropriate
		python program to execute the script & 2. the apptRemind_Text.py script directory). This is the file which
		is triggered by the Task Scheduler to run everyday at a specified time.

		*references_Text.py this file contains all of the variables that can potentially change in the future cobbled
		together in one file for easy editing of the program. Descriptions of what each variable does and where it can
		be found in the apptRemind_Text.py script are described in this file.

		*requirements.txt contains all of the downloaded packages with version control necessary for the program to run.
		To install this program on another machine, download 'cmd console emulator' (i.e. a program that simulates the
		command line prompt except that it accepts linux commands) and enter "cd the\new\directory\location" then "pip
		install -r requirements.txt". This will automatically download all of the dependencies.

		*Development this file that contains the first version of the automated appointment reminder program. Initially,
		the Centricity reports were going to be modified to include the patient's preferred language and would have two
		phone number fields (one for cell and one for home phone). The idea was to send language appropriate reminders
		as a text message to cell phones and voice messages to home phones. That is also why there are four redcap projects
		associated with this program on only SpaText is used.

	Centricity_Reports is a directory created to receive centricity reports that are automatically created and sent to this file
	by an automated system run by the Centricity application owner who is Chad Waggoner at the time of creation. Each day two new
	files should appear in this folder around 6 am. One file for coporate screenings and another for outreach screenings.

	Event_Locations is a directory created to contain the master schedule for all upcoming screenings that are held on the big
	pink bus.

Installation:

	First install anaconda which is a python distribution with additional features. To see where the previous version was downloaded,
	check the Program_Start batch file. That will tell if it was downloaded under the users directory or a universal directory for
	all users. You want to make sure you are downloading python version 3.6.

	Next, double click the apptRemind_Text.py file. You should get a window that asks which program you want to run the file with.
	Select the python program in the anaconda repository which you just downloaded.

	Install packages listed in the requirements.txt file. Directions are given above in the file description.

	Set up Task Scheduler to trigger the Program_start.bat file at a specified time each day. You want to make sure you provide
	a time which is well after when the Centricity reports are generated. Currently that occurs around 6 am. Therefore, the
	current task scheduler set up is set to run the program at 6 pm to avoid any issues of running while the machine is being
	used for other purposes.

	Run a test and check Redcap under 'manage survey participants' then 'survey invitation log' to make sure all patients are set
	to receive a notification at 12 noon the day following when their records are loaded into the system.

