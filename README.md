# PWD-Manifest-Automation-Solutions

Project Description
------------------------------------------------------------------------------------------------------------------------------------------------------------
This program was made to automate and facilitate the creation of daily installation manifest files for Phoenix Window & Door.
An installation manifest is a document for each of PWD's installation crews indicating information about the crew's daily 
workorders such as the builder they will be installing for, which particular subdivision of that builder they are working in, the 
particular lot they will be at, and the number of windows and doors they will be installing as well as a section for notes regarding each order.
This particular case of manifest files are created as excel spreadsheets.

How To Use
------------------------------------------------------------------------------------------------------------------------------------------------------------
- By default, the Manifest Creator will load up today's manifests. 

- In the date entry box, if you provide a date and click the 'Load Manifests' button, it will attempt to find and load manifests of the given date following 
the folder structure to the appropriate month and year, and load all manifests for the given date if found. If no manifests are found for the given date, it 
will prompt you, asking if you want to create new ones. 

- Alternatively, you can simply create a new batch of manifests for a given date by providing said date and clicking the update manifests button. The program 
always creates new files, therefore if none exist for a given date, new ones are made. If existing ones are found, the contents of those are loaded into 
the program. Changing the contents and keeping the date the same then chosing to update manifests will overwrite existing ones.

- The 'Add New Manifest' button inserts a blank new manifest with 6 empty work order fields above the current manifests.

- Manifests with a blank lead name will not be saved to a spreadsheet.

- The 'Erase Entries' button will delete all text in every entry for the given manifest, excluding the Lead name and Crew member names.

- The 'Print Manifests' button will print all manifests to the machine's default printer for the given date. This function will exclude any spreadsheet
manifest files which do not have the builder field for the first work order filled it as it will assume the manifest is incomplete.


Installation Notes
------------------------------------------------------------------------------------------------------------------------------------------------------------
This program was written in python 3.10.7 64-bit (microsoft store version). 

- An installation of python's openpyxl library is required.

1. 'main.py' - Written using Python's standard GUI module: 'Tkinter'. Run this to start the application.

2. 'ManifestGenerator.py' - Mainly contains program logic. This file is also runnable, however its functions are limited to creating new manifest files and 
printing all manifest files for a given date, determined by a timedelta equation of x amount of days ahead of the current day.

3. 'ManifestLoader.py' - Mainly contains program logic, however running this file will instead import 'main.py', triggering that file to run instead.


Other Important Project Dependencies
------------------------------------------------------------------------------------------------------------------------------------------------------------
This program is designed to work with a very specific file structure as initialized within the 'ManifestGenerator.py' file. The 'unf_manifests_filepath' variable
and the 'get_manifests_filepath' method are used to determine where the manifest spreadsheets are read from and stored.


Author
------------------------------------------------------------------------------------------------------------------------------------------------------------
Diego Suarez

Github: https://github.com/dsuarez80/PWD-Manifest-Automation-Solutions

Email: dsuarez8085@gmail.com