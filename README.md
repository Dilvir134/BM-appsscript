# BM-appsscript
In this project I am using apps script to manage Daily Cash Entries, and manage Customers of BM Enterprisees.

The main apps used are google sheets and google forms.

Both projects are stored in the main [BM Folder](https://drive.google.com/drive/u/0/folders/1ntBurl3x5MqJyPYYiwdZ-6umh35v5Dib)

[Daily Entries](https://docs.google.com/spreadsheets/d/1tQBGobBxjISfrwqO2xFE9CeyQ3NI7Cfbw_tmGBgjhRI) are stored in a spreadsheet, each day format is duplicated and renamed to the date of the day by the user. Type options are given if there are finance entries or balance entries. Two default finance options are given:
1. A.M.C.L (Finance)
2. P.H.F Leasing (Finance)

To add another financer we can simply select the whole first column, de-select the header and click data validation, to the list of items add the financer name in this format : "{{ name }} (Finance)".

The format sheet has columns of **Status** and **Balance**, these are to be used as such:

Let's say we have an entry where someone leaves a balance of Rs. 500. We would select balance from type and enter the details, we would put any cash in if entered and write the amount of balance he leaves in the balance column. We then click automation tools and move balance to individual sheet. This is also done with the finance entries, but instead of entering the amount of balance, the amount that is picked is from the **Cash in** column.

When a finance entry / balance is paid, the updation of the status is done as such:
On the day of payment we will store a new entry of the payment with certain details, then we can find the entry and select the link from the individual sheet to go to the certain entry and select the paid option from the status column, then right click the cell and insert link, then from **Sheets and Named Ranges** then click select a range of cells to link. Navigate to the entry where the payment is made and select the cell. Once done move the balance / finance to its individual sheet agian.

This system has some limitations however, for example when naming a new sheet you have to make sure after duplicating the format you have to name it in the specific format of dd/mm/yyyy. If it is even slightly different e.g d/mm/yyyy the system will not work as intended.
