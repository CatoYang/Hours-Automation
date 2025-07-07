# Hours Automation Project
This is a repository of code for automation of hours processing.  
The hours processed are for RSAF F15 propulsion flight and split into two branchs. SG and PCV  
The code listed here is mainly for PCV, SG hours is a fork and due to lack of processes have different code and is slightly deprecated.  
There is also no confidential data here. As the code only post processes it.

# Python Legacy Project
The main project mainly focuses on VBA code but a python project is included that works on sheet formats pre 2025.  
Due to local machines unable to run python, users will have to set up a executable or docker for the legacy code as well has handle the dependencies.  
As the old python project is not in use anymore, seek documentation found in the source sheets from jan 2025 to may 2025.  
The code is found in process.py under _src  
A version of this project could also exist as a executable on local machines, left deprecated after the jump to VBA  

# VBA project overview
The project is 100% VBA code and revolves around 3 main macros and one module from which functions are listed

All of the _Update macros have code to timestamp when they are run in Daily_Hr  
The .cls files are a way for buttons to be implemented for older versions of excel
```
Hours
├── AFH_Update          # Checks and updates AFH values then runs SyncFillColours
├── TTC_Update          # Checks Values and Time then overwrites values
├── LRU_Update          # Checks Values then overwrites and syncs hours to start clocking difference
├── Function_List       # A place to put all minor functions and sub
│   └── SheetExists     # Function to check if sheets exist before running
│   └── SyncFillColors  # Syncs fill colour for visual indication in the main page
├── Sheet1.cls
├── Sheet2.cls
├── Sheet10.cls
```
# Using the code
Its easier and faster to copy code directly into the VBA editor if changes are made here and you need to update.  
Because the code was exported, avoid copying the meta-data and attributes in the files when updating it.  

# Explaination of Macros
This is really just a quality of life update, practically users should know how it inately works so you can manually copypaste (like how it was originally done)  
These macros also heavily use the processor and ram so take note to close processes so excel can ultilise it (the macros have subpar performance on old laptops and machiens under load)  
The following are a brief explaination of the macros.

### AFH_Update
A simple macro that checks and updates the AFH  
```
Checks target sheet exist with SheetExists()  
Reset the previous highlights  
Colours column 6 yellow to signify editable column  
Reads the values to be compared  
Depending on the values fills the cell with colour  
- If higher, colours green  
- If higher by 6 , colours purple and fires a caution  
- If Lower, colours red and fires a warning  
- If equal, nothing changes and the cell is yellow  
Runs SyncFillColors to mirror changes onto the main page  
Timestamps main page and sets format  
executes a message box detailing completion  
```

### TTC_Update
This Macro checks data in TTC and overwrites data in Engines if Download time is earlier and EOT is higher.  

```
Checks target sheet exist with SheetExists()  
Loads data into a dictionary  
Resets fill colour  
For every engine in Engines, checks if it exists in TTC  
If yes, Checks if EOT and Download time is higher/earlier  
If yes, overwrites the old data 
Highlights rows if there was a overwrite
Timestamps the overwrite (so users can easier see when was the last download)  
Timestamps the date/time of execution in Daily_Hr  
Executes a message box detailing completion  
```
### LRU_Update
This Macro is for parts tracking for 'line replaceable units'. It takes PSN data as raw data.  
As it takes data from PSN, ETDS must have updated data before exporting it to input into the workbook.  
This doesnt have to be executed often, only when components are removed and installed do we require to input for tracking.  
The macro automatically updates if ETDS is updated properly  

```
Checks target sheet exist with SheetExists()  
Loads PSN data into a dictionary  
Loads LRU data into a dictionary  
Resets fill colour in LRU  
Checks LRU for every component serial no if it exists in PSN  
If yes, checks for a EOT (this is actually COT) increase  
Overwrites old data  
Highlights and timestamps (the timestamp is in a hidden row)  
Syncs the hours to the current EOT of the engine so it can begin tracking  
Resets fill colour in PSN  
Checks for Serial numbers not present in LRU  
Highlights those serial numbers (this is a test to see any components that need tracking that are missing)  
Timestamps the date/time of execution in Daily_Hr   
Executes a message box detailing completion  
```

### Function_List
This is a module that stores code for minor functions that are not big enough on its own  

SheetExists()  is for checking sheets exist.  
Its mainly created so users will not break the system by casually renaming sheets without prior knowledgge.  
If the next user capable of understanding its purpose comes along please feel free to change the name designations  

SyncFillColors() is for visual update for main page  
This is a aesthetic upgrade that fires when AFH_update is run and mirrors the highlights onto the main page  
The purpose is for FLC to have a better visual indicator of changes.  

### SQN_Report
This exists only in SG  
Reason being we only need to send out specific information so i created a sheet that only pulls that information (SQN_HRS)  
Then we needed to send out LRU data so we used Encik Lew's sheet. (a deal was struck for external monitoring)  

Then i created PSU, LRU sheets to standardise and automate updating of the sheet.  
The two sheets are superior in accounting and maintenance as the old sheets required you to manually edit all values and there are no instructions for pulling data.  

There are a few reasons why the encik lew's sheet was and still is used over my system as of june 2025;  
- Sending out the entire LRU sheet would invite questioning regarding the other unimportant data (image control by management)  
I offered to include a critical subset of the data so it alights with what was shown but it was not adopted.  
- The old system is maintained and new users have failed to maintain my system in SG  

This is projected to lead to a split between use cases for frankly poor reasons.  
Regardless, this is what the macro does  
```
Creates a new excel sheet with 'Hours for SQN' && Date  
Copies over SQN HRS sheet  
Copies over encik lew's LRU sheet  
Cleans up  
```
In the Future there will be a problem of updating this if the old system is updated to the new one.

# Development Notes
Python is not available on local machines. Don't Bother with trying to use it.  
Because there are many machines that have different specifications, there are a few coding differences when compared to modern code.  
1. we are unable to call the entire macro without using thisworkbook.thisworkbook.module.
2. Certain data validations and functions are not modernised. e.g. we use vlookup instead of xlookup
3. Xlookup is not available on 2019 excel (Our version is not the enterprise version?!)

The earliest version of Excel backtested is Excel 2010. The VBA macros do not work on Excel 2007.
Most personnel are unable to comprehend VBA, but some NSmen are actually savants with it (NSman Benson could automate ES with VBA code!!). Talk local personnel and seek their strengths.

# Contribution
- Fork the repo
- Create your feature branch
- Commit your changes
- Push to the branch
- Open a Pull Request

I will perform code reviews for this project as long as i am able.  
Alternatively contact me directly and we can sort out the changes together  

# Contact
Message me directly on Whatsapp or Telegram.  
Alternatively contact via github or my email
