# Fixing and Converting Voucher Wise Deposit Report to xlsx
 Automation process designed for an institute

"Voucher Wise Report" daily uploads on website for students query. Website is not the part of institute's EMS so they download that report daily and covert that csv report to our required format. Voucher Wise Report downloaded has 2 issues.

1- Receipts Column is Blank (So we'll read the narration and place the number in the column)

2- Downloaded Format is csv and CSV format has date issue. Sometimes it shows year in the first section YYYY-MM-DD and sometimes DD-MM-YYYY

Wroking

Download the report from EMS and place it on folder and make sure that it has its default name "voucher_wise_deposit.csv"

Double Click on the application it will read the csv file convert it to xlsx file and fix errors.

Then open the output file in the xlsx format save that file again in ods by the name of receipts_grand.ods file format (that name because website database table name is receipts_Grand changing that name will create another table on the sql side and website wont fetch values from new table) and upload it on the sql of the website.
