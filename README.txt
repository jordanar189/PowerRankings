- This project creates a fantasy football power rankings graphic based on a file of polls. In its current form, it is specific to the league of the creator of this project, but minimal revisions can make it applicable to any 10-team fantasy league. Read below for information on each file:


- *run.py: This is the file that is used to execute the code and calls on functions from all other Python files

- *create_df.py: This Python file creates a dataframe to hold point information for each team and also assigns the point values to the respective teams in the dataframe

- *r_graph.py: Although this is a Python file, it uses mostly R. This file will create the graph to be attached to the final graphic

- *create_graphic.py: This file has 3 functions. One to add a "first place votes" indicator, the next one sets the NFL week, and another creates the final graphic


---Fonts folder---
- *URW Grotesk Regular.ttf: This is the font used on the final graphic


---Pictures folder---
- *graphic.png: This is the background for the final graphic

- GraphicFinal.png: This is what the final graphic will be saved and labeled as, after the code finishes running. Currently not included in folder until after code runs

- graph.png: This is the graph created by r_graph.py. Like GraphicFinal.png, this will not appear until after the code is run once


---Tables folder---
- *Power Rankings.xlsx: This is where the polls are stored. You can change the items being ranked by updating each team name within each poll to whatever items you like. Each item must be seperated with ", ". There can only be an exact number of 10 items in each poll. There can be an unlimited number of polls (each row is a new poll), but each poll must consist of the same 10 items, however the order of those 10 items can differ poll by poll. (Despite the asterik, this file can be updated in the manners described here, however it cannot be deleted)

- points_calculated.csv: Like others, this will not appear until the code has been run at least once. This is the dataframe that was collected in create_df.py. Helpful for checking results


***Files with a asterik preceeding the file name can not be adjusted or deleted or else the code will not run properly***
