# Hours Automation Project

This is a repository of code for automation of hours processing.
The hours processed are for RSAF F15 propulsion flight and split into two branchs. SG and PCV
The code listed here is mainly for PCV, SG hours is a fork and due to lack of processes have different code and is slightly deprecated.

The project mainly focuses on VBA code but a python project is included that works on sheet formats pre 2025.
Due to local machines unable to run python, users will have to set up a executable or docker for the legacy code as well has handle the dependencies.
As the old python project is not in use anymore, it will not be explained hereforth.
The code is found in process.py under _src

# VBA project overview
The project is 100% VBA code and revolves around 3 main macros and one module from which functions are listed

All of the _Update macros have code to timestamp when they are run in Daily_Hr
The .cls files are for the 

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

# Using the code
Its easier and faster to copy code directly into the VBA editor if changes are made here and you need to update.

# Explaination of Macros
to do


# Development Notes
Python is not available on local machines. Don't Bother with trying to use it.
Because there are many machines that have different specifications, there are a few coding differences when compared to modern code
- we are unable to call the entire macro without using thisworkbook.thisworkbook.module.
- Certain data validations and functions are not modernised. e.g. we use vlookup instead of xlookup
- Xlookup is not available on 2019 excel (Our version is not the enterprise version?!)
The earliest version of Excel backtested is Excel 2010. The VBA macros do not work on Excel 2007.
Certain individuals are offended that this project exists. Becareful and employ image control when working with NCOs.
Most personnel are unable to comprehend VBA, but some NSmen are actually savants with it (NSman Benson could automate ES with VBA code!!). Talk to people and seek their strengths.

# Contribution
- Fork the repo
- Create your feature branch
- Commit your changes
- Push to the branch
- Open a Pull Request
I will perform code reviews for this project as long as i am able.
Alternatively just send me the file to edit directly if you do not know git.

# Contact
Yang Peng - yangcymbus@gmail.com
