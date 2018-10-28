# Survey-Feedback-Machine
Also in README.docx
Members: 
  Ang Yock Kang (1) 2I
  Yin Lye Ting (31) 2B
  
Attempted Problem:
  CS Feedback Survey (meant to work dynamically with other surveys too)
  
How to Utilise Program:
  Packages:
  Make sure that these packages are installed before running the program

  1.	Sumy
  2.	Numpy
  3.	Matplotlib
  4.	Docx

  Files:
  
  •	Responses.csv
  
    o	This is where the data is input by the user
  
    o	The user is recommended to copy and paste the data inside this file
  
    o	Sample response set included
  
  •	Config.csv
  
    o	This is where the user can enter the configurations for different surveys and functions utilized by the program
  
    o	Appropriate sample config file included
  
  •	Program.py
    
    o	This is the main program. To run it, one can simply double click it, and there is no need to enter the IDLE text editor
  
  •	Ouput.docx
  
    o	This word document is where an analysis report of the survey is written to by the program
  
    o	Do note that previous reports won’t be automatically saved, so the user is required to manually save it as another file before re-using with different configurations.
  
  •	Condensethis.txt, Output.txt, Error.txt
  
    o	These are side files required by the program, with no need for user interaction.

  (Example configs in file)
  
  In responses.csv, copy-paste the survey data
  The first row is dedicated to all the questions,
  Below that is the data

  In config.csv, enter the configuration for your questions
  Every row corresponds to your question number
  In the first column, enter
    
    “Admin” for administrative questions,
    
    “Num” for numerical questions,
    
    “Categorical” for categorical questions,
    
    “Qualitative” for qualitative (open ended) questions

  For Admin and Qualitative questions, nothing else must be entered
  For numerical questions, please enter the first and last numbers on columns 2 and 3 respectively.
  For Categorical Questions, please put all your categories to the right of the cell in “Categories”, please dedicate 1 cell for each category

  There are supplementary functions, they are:
  •	Graph Type(Either Pie or Bar)
  
    o	Pie or Bar: returns charts in that format for Numerical and Categorical questions
  
  •	Executive Analysis
  
    o	ONLY ONCE
  
    o	There are two options, CORRELATION and DIFFERENCE, put on the cell below “Executive Analysis”
  
    o	Questions you want to correlate to the right, in separate boxes
  
    o	Only uses data from ‘Num’ questions
  •	Correlation
  
    o	This function analyses the trends between the responses of 2 numerical questions
  
    o	Enter the numbers of the question that has the independent variable (cause/x factor) and the dependent variable(effect/y factor) beside the Correlation file
  
    o	Returns a scatterplot graph with a linear trendline and a pearson correlation coefficient derived statement that describes the strength of the correlation between the data and the trendline
  
    o	Example: If i am doing a research on the effects of time spent on reading books on grades, and my survey is done on students, I can put the question that asks for the time spent as the x factor and the grades as the y factor, and the relationship between the two
  •	Difference
  
    1.	Calculates the difference for every respondent between two questions, analyses it holistically and generalises it in 1 statement
  
    2.	Example: If qn1 is on how the student likes the cs course, and qn2 is whether they liked a certain part of the cs course, this function can be used to analyse whether qn2 is ranked higher, lower, or similarly to qn1 so the teacher can adjust accordingly
  •	Summary for perception of program?(Y or N)
  
    o	If Y, it will provide a summary
  
    o	If N or inappropriate, it will return no summary
  •	Demand Question Number
  
    o	Input number you want to count demand
  
    o	Input options to add to demand to the right, one option per box
  Note: functions to be applied to the second column


  E.g.
  Config.csv

     |Admin			|                                           |           |       |
     |Num	      |1                                          |10         |       |
     |Num	      |1                                          |5          |       |
     |Num	      |1                                          |7          |       |
     |Categorical|Yes                                       |No         |Perhaps|
     |Qualitative|                                          |           |       |		
     |           |Graph Type (Either Pie or Bar)		        |           |       |
     |           |Bar (nil or inappropriate will return Bar)|           |       |
     |           |Executive Analysis		                    |           |       |
     |           |Correlation                               |2          |   3   |
     |           |Summary for perception of program?(Y or N)|	         |       |	
     |           |Y	                                        |           |       |
     |           |Demand Question Number		                |           |       |
     |           |4(Blank will not give expected demand)    |Yes(option)|       |	


  How the program works
  In a nutshell, the program reads the data from the configurations file and the data file and stores the data in lists. Then, the program analyses the data based on the configurations in the config file, then writes an analysis report to the output word document

  Functionality:
  
  All suggested functionality in rudimentary and intermediate sections
  Sumy and Matplotlib
  Configuration file(config.csv)
  Summary (consolidated picture of how students perceive the cs course+demand for cep)
  Executive Summary (Either correlation or difference)
  .docx implementation, writes text and graphs alike 
