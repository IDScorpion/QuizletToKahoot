import os
import openpyxl

if os.path.exists("uploads") is False:
    os.mkdir("uploads")

uploaded = False
while uploaded is False:
    fileName = input("Please upload a .csv file of your quizlet to the folder uploads. Use , as the term-def delimiter, and new line between rows. Name the file anything with no spaces or special chars. Press enter the name of your file. ")

    partitioned = fileName.partition(".")
    if partitioned[1] == "":
        fileName = f"{partitioned[0]}.csv"
    if os.path.isfile(f"uploads/{fileName}") is True:
        uploaded = True
        print("\nAwesome! I found your file! Let's make a Kahoot!\n")
    else:
        print(f"\nI couldn't find file {fileName} . Please try again. \n")

kahootTemplateWorkbook = openpyxl.load_workbook("kahootTemplate.xlsx")
kahootTemplateSheet = kahootTemplateWorkbook['Sheet1']

# B9 is Q1 start 
kahootTemplateSheet["B9"].value = "Testing"

kahootTemplateWorkbook.save("test.xlsx")