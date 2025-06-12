import pandas as pd
import sys
import os
import time
import logging
import traceback

## Times the start of the script for efficiency purposes
start_time = time.time()
## Input date into the
current_time = time.localtime()
formatted_date = time.strftime("%d-%m-%Y", current_time)
formatted_time = time.strftime("%H:%M:%S", current_time)

# Configure logging
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("debug.log"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# Function to handle uncaught exceptions and log them with traceback
def log_uncaught_exceptions(exc_type, exc_value, exc_traceback):
    if issubclass(exc_type, KeyboardInterrupt):
        # Ignore keyboard interrupts to allow clean exit
        sys.__excepthook__(exc_type, exc_value, exc_traceback)
        return
    logger.error("Uncaught exception", exc_info=(exc_type, exc_value, exc_traceback))

# Assign the custom exception handler to sys.excepthook
sys.excepthook = log_uncaught_exceptions

with open('Report.txt', 'w') as file:
    sys.stdout = file
    ##Primer to User
    print('=============================================================== TTC copy, Tests and Verifications Script ===============================================================')
    print('For changes and troubleshooting please contact ME2 Yang Peng')
    print('Email: yangcymbus@gmail.com, *GitHub repository pending implementation')
    print('Links to Versions https://drive.google.com/drive/folders/17y9LJ3d8apgNYHIt-F8TKDw527QuyJ41?usp=sharing')
    print('Version 0.4')
    print('Executable compiled on 10/1/202')
    print("Script Executed on: ", formatted_date, formatted_time)
    print()

    ## Gets the longest file name so we dont have to set a static name. We presume the longest name is the Engine Monitoring Report
    Report = max((f for f in os.listdir('.') if os.path.isfile(f)), key=len, default=None)

    # Kills the Script if we dont have a file called TTC.xlsx
    if 'TTC.csv' not in os.listdir('.'):
        print('TTC file not found, ensure the file is in the folder and is named "TTC"')
        sys.exit()

    ## all our files are named TTC from CETADS
    df = pd.read_csv("TTC.csv")

    ## Prints the File name we are processing for confirmation
    print(f'Modifying file: {Report}')
    ## Number of Rows of Data that needs to be processed
    nrows= df.shape[0]
    print(f'Number of Engine Rows: {nrows}')
    ## dataframes start at 0, we need to compensate
    nrows-=1
    ## N = row number in the dataframe starting with 0
    n=0
    print()
    print('=================================================================== ESN & EOT Verification and Copy ===================================================================')
    ##Formating output table for the loop, ANSI escape codes aren't being processed by this environment, we cannot underline this 
    print(f"{'Row':<5}{'A/C':<5}{'|'}{'Original ESN':<13}{'Original EOT':<13}{'L/R':<5}{'|'}{'Input ESN':<13}{'Input EOT':<13}{'L/R':<5}{'|'}{'Remarks'}")

    while n <= nrows:
        ## For attaining the correct ASN from a data frame
        #Pulls A/C number
        ASNO = df.loc[n ,'ASN']
        #Strips A/C number to last 2 Digits
        ASNO2= ASNO[-2:]
        #Adds 'N' to the front 
        ASNO3= 'SG' + ASNO2
        ##Reads the file we are supposed to edit and convert to dataframe
        df2 = pd.read_excel(Report , sheet_name= ASNO3)

        ## EOT comparisons before copy
        NEW= df.loc[n, 'EOT']
        ## Identify which Engine is it
        ENGNO= df.loc[n, 'ESN']
        ##Stripping Dataframe into a single line with the exact engine we need
        fdf1 = df.loc[[n]]
        #finding the row for the data set then setting it to write later, this is based on the TTC engine position, we have to invert it due legacy formatting
        #we put .any() because sometimes we get int vs str issues
        if (fdf1['Location'] == 2).any():
                SR=1
        else: 
                SR=2
        #We Subtract 1 from SRX because pandas considers the first row to be 0
        SRX = SR - 1

        while len(df2) < 2:
        # Append an empty row (NaN values) to make sure the DataFrame has 2 rows because if theres a missing engine it will break
            # Create a new row with None values (empty row)
            empty_row = pd.Series([None] * len(df2.columns), index=df2.columns)
            # Concatenate the empty row with the existing DataFrame
            df2 = pd.concat([df2, empty_row.to_frame().T], ignore_index=True)

        #Increase row value by one due to first value being zero
        nr = n+1

        if df2.loc[SRX, 'ESN'] == None:
            print(f"{nr:<5}{ASNO3:<5}{'|'}{ 'NA':<13}{'NA':<13}{'NA':<5}{'|'}{ENGNO:<13.0f}{fdf1['EOT'].values[0]:<13}{loc1:<5.0f}{'|'}", end="")
            print('No Engine detected, Copying data...', end="")
            with pd.ExcelWriter(Report , mode="a", if_sheet_exists="overlay") as writer:
                        fdf1.to_excel(writer, sheet_name= ASNO3, header=False, index=False, index_label=ENGNO, startrow=SR)
            print('Data is successfully written')
            
        else:
            ENGNO2= df2.loc[SRX, 'ESN']
            
            ## Obtaining dataframe of the exact row we want to copy
            fdf2 = df2[df2['ESN'] == ENGNO2]
            fdf1.reset_index(drop=True, inplace=True)
            fdf2.reset_index(drop=True, inplace=True)
            ## converts values for Boolean comparison
            comparison_results = (fdf1['EOT'].values > fdf2['EOT'].values)

            ##Setting values for the Engine position of both old and new inputs
            loc1= fdf1['Location'].values[0]
            loc2= fdf2['Location'].values[0]

            print(f"{nr:<5}{ASNO3:<5}{'|'}{ ENGNO2:<13.0f}{fdf2['EOT'].values[0]:<13}{loc2:<5.0f}{'|'}{ENGNO:<13.0f}{fdf1['EOT'].values[0]:<13}{loc1:<5.0f}{'|'}", end="")

            if ENGNO == ENGNO2:
                print('Engines are the same, ', end="")
                if comparison_results: 
                    print('TTC data is higher, Executing copy... ', end="")
                    with pd.ExcelWriter(Report , mode="a", if_sheet_exists="overlay") as writer:
                        fdf1.to_excel(writer, sheet_name= ASNO3, header=False, index=False, index_label=ENGNO, startrow=SR)
                    print('Data is successfully overwritten.')
                else: 
                    print("TTC data is equal or lower, skipping data input")
            else: 
                print('Engines are not the same, skipping input')
            
        ## Increases n by one, repeat the loop
        n += 1
    print('============================================================================== Last Entry =============================================================================')
    


    end_time = time.time()
    elapsed_time = end_time - start_time
    print()
    print(f"Execution time: {elapsed_time:.4f} seconds")

    ## To do, 200,400,800 hrly unit test
    ## PTO Shaft due time test
    ## LRU due test
    ## pings for impending due
    ## 

    ## To build the executable, assuming pyinstaller is already installed, run the command below. add --noconsole if you dont want a console window to appear
    # pyinstaller --onefile process.py 
    # If for some reason pyinstaller isnt detected, do 'pip install pyinstaller'

    sys.stdout = sys.__stdout__