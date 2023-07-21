# catia2excel
Import and Export 3D Shapes from CATIA to Excel

This project is to help facilitate an integrated workflow from CATIA to Excel.

**export_excel.bas** : Export Points from CATIA to Excel ( Excel workbook needs to be opened concurrently with CATIA before running the script )

Steps to run this script : 
1. Download this repository
2. From within CATIA, launch VBA ( Alt + F11 ) 
3. Create a library ( VBA IDE should open up at this stage )
4. File > Import File ; Navigate to the downloaded repository folder and select the script
5. The file should be displayed on the project window on the left panel, double click on filename to open it
6. Add additional references :
  - Tools > References > Select "Microsoft Excel" Object Library ( Title is usually of the form "Microsoft Excel xx Object Library" )
7. Open Excel, create a new spreadsheet, and select starting cell 
8. Run this script ( F5 )

