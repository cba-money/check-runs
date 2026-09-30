# Check Runs App
This program is built to reconcile two datasets, a check register (which registers all checks written) and a check run (which includes all checks cashed within a certain time period).

## How it Works
The application checks all of the cashed checks against the register to audit the amount and the number of times they were cashed. It also checks for and marks voided checks. After processing, the app returns three files: A modified check register (which has the "Date Cashed" column marked with each date cashed for each check respectively), a modified check runs (which has each check marked either blue or yellow, depending on if it incured a discrepancy or not), and a discrepancies file (which lists all discrepancies found in the run).

After running up to the current date, the ouput check register can be used to quickly check if and when a certain check was cashed. The process also identifies discrepancies which can be fixed through the bank.

## Development
Make sure Node.js and NPM are installed on your machine. Clone the repository and open a command prompt in the root of the direcotry.

Run:
```sh
npm install
```

To install dependencies, then run:

```sh
npm run dev
```

## Installing
To install for Windows, download the latest .exe from the Releases. Run the .exe and the application should automatically install to the machine.

## Usage
To use, open the app. Click on "Choose Register" and select the check register. Next, click on "Choose Run" and select the check run file. (Note: both files must be .xlsx or it may not work as expected).

Press the "Start Process" button and processing should begin.

Processed files will appear in an "exports" directory (in the same folder as the installed app).



