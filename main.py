import csv
import os
import random

import openpyxl


class Question:
    def __init__(self):
        self.question_text = None
        self.answers = {
            "1": None,
            "2": None,
            "3": None,
            "4": None
        }
        self.time_limit = None
        self.correct_answer = None

    def add_answer(self, answer):
        amount_answers = len(self.answers)
        if amount_answers == 4:
            if self.answers["1"] is None:
                self.answers.update({"1": answer})
            elif self.answers["2"] is None:
                self.answers.update({"2": answer})
            elif self.answers["3"] is None:
                self.answers.update({"3": answer})
            elif self.answers["4"] is None:
                self.answers.update({"4": answer})
        elif amount_answers < 4:
            next_key = str(int(amount_answers) + 1)
            if int(next_key) <= 4:
                self.answers.update({next_key: answer})
            else:
                return False

    def check_requirements(self):
        if self.question_text is not None:
            if self.answers["1"] is not None and self.answers["2"] is not None:
                if self.time_limit is not None:
                    if self.correct_answer is not None:
                        return True

        return False

    def cleanup_answers(self):
        if self.answers["3"] is None:
            del self.answers["3"]
        if self.answers["4"] is None:
            del self.answers["4"]


current_row = 9

kahoot_template_workbook = openpyxl.load_workbook("kahootTemplate.xlsx")
kahoot_template_sheet = kahoot_template_workbook['Sheet1']


def append_to_sheet(question, current_row=current_row):  # Assumes a Question Object

    layout = {
        f"B{current_row}": question.question_text,
        f"C{current_row}": question.answers["1"],
        f"D{current_row}": question.answers["2"],
        f"E{current_row}": question.answers["3"],
        f"F{current_row}": question.answers["4"],
        f"G{current_row}": question.time_limit,
        f"H{current_row}": int(question.correct_answer)
    }
    for key in layout.keys():
        kahoot_template_sheet[key].value = layout[key]
    current_row += 1


def save_sheet():
    kahoot_template_workbook.save('output.xlsx')


if os.path.exists("uploads") is False:
    os.mkdir("uploads")

uploaded = False
while uploaded is False:
    fileName = input(
        "Please upload a .csv file of your Quizlet to the folder uploads. Use , as the term-def delimiter, and new "
        "line between rows. Name the file anything with no spaces or special chars. Press enter the name of your file. "
    )

    partitioned = fileName.partition(".")
    if partitioned[1] == "":
        fileName = f"{partitioned[0]}.csv"
    fileLocation = f"uploads/{fileName}"
    if os.path.isfile(fileLocation) is True:
        uploaded = True
        print("\nAwesome! I found your file! Let's make a Kahoot!\n")
    else:
        print(f"\nI couldn't find file {fileName} . Please try again. \n")

# B9 is Q1 start
"""
kahootTemplateSheet["B9"].value = "Testing"

kahootTemplateWorkbook.save("test.xlsx")
"""

questionLengthChoices = (5, 10, 20, 30, 60, 90, 120, 240)
allowedChoice = False

while allowedChoice is False:
    print("How long do you want per question, in seconds?")
    questionLength = int(input(f"Your choices are: {questionLengthChoices} "))
    if questionLength in questionLengthChoices:
        allowedChoice = True
        confirmChoice = input(f"\nEach question will be {questionLength} seconds. Is this correct? (Y/N) ")
        validConfirm = False
        while validConfirm is False:
            if confirmChoice.lower() == "y":
                print('\nAwesome!\n')
                validConfirm = True
            elif confirmChoice.lower() == "n":
                print("\nOK. I'll go back a step.\n")
                validConfirm = False
                confirmChoice = input(f"\nEach question will be {questionLength} seconds. Is this correct? (Y/N) ")
            else:
                print("\nIt seems you entered something other than Y or N. We'll try again.\n")
                validConfirm = False
                confirmChoice = input(f"\nEach question will be {questionLength} seconds. Is this correct? (Y/N) ")
    else:
        print(f"\n{questionLength} isn't a valid length. We'll try again.\n")

print("I've got all the information I need from you. I'll let you know when I'm done.")

# Begin building template


quizletData = {}

with open(fileLocation, 'r') as file:
    readable = csv.reader(file)
    for row in readable:
        quizletData.update({row[0]: row[1]})

position = 0
all_keys = []
all_values = []

for key in quizletData.keys():
    all_keys.append(key)
for value in quizletData.values():
    all_values.append(value)
while position < len(all_keys):
    current_question = Question()
    current_question.question_text = all_keys[position]

    answer_list = [all_values[position]]
    for i in range(2, 5):
        randomChoice = random.choice(all_values)
        while randomChoice in answer_list:
            randomChoice = random.choice(all_values)
        answer_list.append(randomChoice)

    random.shuffle(answer_list)
    correct_index = answer_list.index(all_values[position])
    current_question.correct_answer = str(correct_index + 1)
    for i in range(1, 5):
        current_question.answers[str(i)] = answer_list[i - 1]

    current_question.time_limit = questionLength

    append_to_sheet(current_question)
    position += 1
save_sheet()
