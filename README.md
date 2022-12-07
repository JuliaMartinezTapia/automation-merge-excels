# Automation of the merge of several excels into one. 

The use case is the following: 

You are an in-house tax proffesional responsible for the consolidated corporate income tax of the group.
You receive the individual corporate income tax calculations in excel files, one per corporation. Then you need to join them in the same tab of the same excel spreadsheet to perform the consolidated calculations. If your group is Spanish, you do this task four times a year (three payments on account and one final calculation at FYE) and you do it manually.

This program allows you to automate this task of merge the individual calculations into a sole excel spreadsheet.

Each time you run the code, the program extracts the individual calculations from the excel spreadsheets and merge them in a newly created spreadsheet named "merged.xlsx" which will include:

- A first column with the concepts of the individual calculations, and 
- second and subsequent columns containing the individual calculations extracted from the excel spreadsheets.

For the program to work properly, the excel spreadsheets to merge need to have the same structure:

- All the excels must contain only one tab with the individual corporate income tax calculation
- the concepts of the calculation must be included in the first column of the tab, and must be equal for all the excels to be merged. 
- therefore, the number of rows must be the same for all the excels to be merged. 

You must bear in mind that the blank spaces among rows will be deleted by the program. Otherwise the merge would not work well.

As an example, I provide nine excel spreadsheets named from "Corp 1.xlxs" to "Corp 9.xlxs", that will be merged into the spreadsheet "merged.xlsx" when the code is run.
Just download the nine excels and the file merge_excels.py to the same directory and run the file merge_excels.py.

Te program will create a new directory "output" within your current directory to store the output excel ("merged.xlsx).

Hope this program is useful for you. Any comment or suggestion will be very good received.

Julia María Martínez Tapia

October 25, 2022
