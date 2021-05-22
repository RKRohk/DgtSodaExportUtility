# DGT Soda Export Utility

This repo has two scripts 
-   exportToXlsx.py
-   importFromXlsx.py

## exportToXlsx

Takes the newest backup file in the folder and exports it to an excel file that you can edit

## importFromXlsx

Takes the edited excel file and converts it into a DGT Soda backup file which you can restore

## Steps

1. Install all requirements from requirements.txt using `pip install -r requirements.txt`
2. Run `python exportToXlsx.py` in the folder containing the DGT Soda backups
3. Edit the new excel file
4. After saving the excel file run `python importFromXlsx.py` to generate a new backup file
5. Import the backup file from the app
6. DONE!