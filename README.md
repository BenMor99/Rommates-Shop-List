# Rommates-Shop-List
This repository is created to manage the project "Roommates Shop List"!
This project aims to provide a beautiful and easy to use space for recording expenses of a dormroom with 3 roommates.
Here there is an .xlsx file as the database.


In phase one, the two main part of the program has been completed: main and reset
reset(), as it's obvious from its name, deletes the items that has been created before and saves a raw and empty xlsx file (except for the headers and column specifiers)
main() function lets you to add an item to the list by recieing your name, information about bought item and the name of who you bought the item for!
you will see an error if:
1. your xlsx file has been unwantedly modified and the original look of the list has been perished
2. you enter a wrong name or misspell it

once you add an item to the list for the first time the program will do summation calculations and add the sum of costs beneaath the column in a cell with different color and as you add more items this cell automatically goes down and down!

in phase two (specifier), I added another module (specifier) including only one function (spec()) to the project.
This module has the duty of calculation of debts and credits and updating the information of payments in a .txt file.
Commiting every change to the Expenses.xlsx (including adding an item or resetting) would trigger the module specifier and message.txt would be updated!
