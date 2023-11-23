To generate multiple Word files with a common letter but unique names and roll numbers from an Excel sheet,
you can use a scripting language like Python. The 'python-docx' library can help create Word documents,
while 'pandas' can assist in reading Excel data.


Certainly! To generate multiple Word files for formal applications with personalized details from an Excel sheet,
you'll need a Word template that contains placeholders for the name and roll number.
Then, a Python script can replace these placeholders with the respective values from the Excel sheet.

#Preparation:
1.Install Required Libraries: Ensure you have installed the required Python libraries, pandas and python-docx.

2.Create a Word Template:
Design your formal application template in a Word document.
Insert placeholders like <<NAME>> and <<ROLL_NO>> at positions where you want to replace the values.

#Python Script:

1.Import Required Libraries:
2.Read Excel Data:
3.Create Output Folder:
4.Load Word Template:
5.Define Font and Size:
6.Iterate Through Excel Data and Generate Documents:

Adjustments:
Replace 'your_excel_file.xlsx' with the path to your Excel file.
Modify 'path_to_your_template.docx' to the path of your Word template.
Adjust column names ('Name' and 'Roll No') to match your Excel data columns.
Customize font and size properties by changing the values of name_font, name_size, roll_no_font, and roll_no_size.
Execution:
Run the script in your Python environment.
It will generate personalized Word documents using the template for each row in the Excel file and save them within a timestamped folder.
