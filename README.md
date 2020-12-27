# Airbnb-API-AutoRoom
## Purpose
This program was build to solve problems that when a group of friends going on a trip,the leading one mostly do servey on hotel rooms or trips,
and contruct a budget data sheet to calculate the payments and feedbacks.


## Funciotn
This program is using airbnb api as room searching program, and Excel as the output format.
Before start running this program, you have to put 'airbnb' file into your python/Lib/site_packages.
Please do not use "pip install airbnb" in the termial because I had make some changes in this package.

After the above movemnet, you can run the program in any complier.
This program will output an Excel xlsx format name "{check_in}{tour_place}.xlsx" and the same file this program is.
The file contain the following three data sheets:

1.Rooms sheet
	Contain every airbnb rooms avaliable currently.
2.Voting sheet
	Contain N airbnb rooms that price per person is below P and is sorted by rated and reviews. (N, P is a parameter that user inputs)
3.Budget Control
	A place for your group to calculate the final payments/feedback to the group.

## Work Sequense

1.Using any software that can run python(.py)

2.Open "AutoCreat_tourExcel.py"

3.Run AutoCreat_tourExcel.py

4.Please input your tour place: (user's tour place)

5.Please input your tourists number: (user's group numbers)

6.Please input check in date (ex: 2020-01-01): (yy-mm-dd)

7.Please input check out date (ex: 2020-01-01): (yy-mm-dd)

8.Enter all groups members name

  Please input tourist 1's name: (First member's name)

  Please input tourist 2's name: (Second member's name)
  ...

9.Program show your request to confirm(T/N)

  Tour Place: 普吉島
  Total Tourist: 8
  Tourist'names are [(Group Members Name)]
  Check in date: (check in date)
  Check out date: (check out date) 
  Confirm your request (Y/N): (Confirm ans)

10.Program show the request rooms of the request.

  After request from airbnb API, (avaliable rooms number) rooms have been found.
  The highest price per person is (highest price) TWD, and the lowest price is (lowest price) TWD

11.Programs asked to input voting data

  What's your highest acceptable price per person: (user's highest acceptable price)
  How many rooms would yo like to show in Voting sheet: (user's wanted vote numbers)

12.Program finish
  Finished. Have a fun trip^^


	
