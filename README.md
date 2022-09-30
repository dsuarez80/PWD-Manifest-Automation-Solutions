# PWD-Manifest-Automation-Solutions

Project Description
------------------------------------------------------------------------------------------------------------------------------------------------------------
This program was made to automate and facilitate the creation of daily installation manifest files for Phoenix Window & Door.
An installation manifest is a document for each of PWD's installation crews indicating information about the crew's daily 
workorders such as the builder they will be installing for, which particular subdivision of that builder they are working in, the 
particular lot they will be at, and the number of windows and doors they will be installing as well as a section for notes regarding each order.
This particular case of manifest files are created as excel spreadsheets.


Installation Notes
------------------------------------------------------------------------------------------------------------------------------------------------------------
This program was written in python 3.10.7 64-bit (microsoft store version). An installation of python's openpyxl library is required.

1. 'main.py' - Written using Python's standard GUI module. This runs the main GUI of the application, by default loading all manifest files for the current 
day, with functionality to load a manifest file for a given day, functionality to create a new manifest file by providing a date of files which do not exist 
yet, updating existing manifests from the provided inputs, as well as the ability to print all manifest files for the date provided.

2. 'ManifestGenerator.py' - This file contains a lot of the core functionality of the project such as creation of excel workbooks and other various essential 
program logic. This file is also runnable, however its functions are limited to creating new manifest files and printing all manifest files for a given 
date, determined by a timedelta equation of x amount of days ahead of the current day.

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