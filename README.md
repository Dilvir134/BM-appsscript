# BM-appsscript
In this project I am using apps script to manage Daily Cash Entries, and manage Customers of BM Enterprisees.

The main apps used are google sheets and google forms.

Both projects are stored in the main [BM Folder](https://drive.google.com/drive/u/0/folders/1ntBurl3x5MqJyPYYiwdZ-6umh35v5Dib)

## Daily Entries Management

[Daily Entries](https://docs.google.com/spreadsheets/d/1tQBGobBxjISfrwqO2xFE9CeyQ3NI7Cfbw_tmGBgjhRI) are stored in a spreadsheet, each day format is duplicated and renamed to the date of the day by the user. Type options are given if there are finance entries or balance entries. Two default finance options are given:
1. A.M.C.L (Finance)
2. P.H.F Leasing (Finance)

To add another financer we can simply select the whole first column, de-select the header and click data validation, to the list of items add the financer name in this format : "{{ name }} (Finance)".

The format sheet has columns of **Status** and **Balance**, these are to be used as such:

Let's say we have an entry where someone leaves a balance of Rs. 500. We would select balance from type and enter the details, we would put any cash in if entered and write the amount of balance he leaves in the balance column. We then click automation tools and move balance to individual sheet. This is also done with the finance entries, but instead of entering the amount of balance, the amount that is picked is from the **Cash in** column.

When a finance entry / balance is paid, the updation of the status is done as such:
On the day of payment we will store a new entry of the payment with certain details, then we can find the entry and select the link from the individual sheet to go to the certain entry and select the paid option from the status column, then right click the cell and insert link, then from **Sheets and Named Ranges** then click select a range of cells to link. Navigate to the entry where the payment is made and select the cell. Once done move the balance / finance to its individual sheet agian.

This system has some limitations however, for example when naming a new sheet you have to make sure after duplicating the format you have to name it in the specific format of dd/mm/yyyy. If it is even slightly different e.g d/mm/yyyy the system will not work as intended.

## Customer Management

Customer Management is done with the help of google forms and google sheets and is stored in [Customer Folder](https://drive.google.com/drive/u/0/folders/1_82IJBJ33PqzsGUR0oRAPvCsBbLoIjjh) in the main BM folder. New customers are added using this [google form](https://docs.google.com/forms/d/e/1FAIpQLSdy7Ya3qZ7oDOQGRUs06aX5r4d9aQuwhjmaeH_l4ooTfJ4IPg/viewform?usp=sf_link) once all details are filled, responses are sent to this [Response Sheet](https://docs.google.com/spreadsheets/d/1cvf6qRzOhVgqX8a3iQkceyf_thKx6OcudB35g2wdbFI). On form submit the script runs and three main functions are run:
1. Generate form 20, 21 files - In the customer folder there is another folder called [Customer Files](https://drive.google.com/drive/u/0/folders/13x_Hr8kFsn-FxBdwVcOd-mNmks99qH3Q) here a folder is created with name of the customer and bill number, templates of form 20 and 21 are copied into the folder and filled with customer's data. If a reponse is edited the folder name will be edited and files will be generated again. After the files are generated the customer's folder link is put in the response sheet.
2. Create an edit url - This function creates an edit url which allows the user to edit their response if anything needs to be changed, also when editing, in the end user is asked if they need to generate form 20, 21 files again, which saves computing power if it wasn't needed.
3. Create customer sheet - In the main response spreadsheet after a response is submitted, a sheet is created with name of customer and bill number, customer details are added as well as the guarantor's details and the sheet link is pasted in the main response sheet. Using the details it creates an account statement format for the specific customer, with columns showing the credit, debit and remaining balance. Another column is also used for setting a reminder date, and when the current date passes the reminder date the row of the customer will turn red as a reminder.
