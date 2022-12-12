def Main():
    #Loading Packages
    import pandas as pd
    from datetime import date, timedelta
    import os
    from tkinter.filedialog import askopenfilename
    import rpy2.robjects as robjects
    from PIL import Image as image
    from PIL import ImageDraw as draw
    from PIL import ImageFont
        
    
    '''Creating Working & Final CSV'''
    #Creating working sheet (ws) to perform calculations and data transformation on and creating CSV
    ws = pd.DataFrame()
    ws.to_csv('.\workingSheet.csv')
    #Using placeholder data to populate columns for final sheet (fs) that will be sent to rStudio
    startingdata = [[0,0,0,0,0,0,0,0,0,0],[0,0,0,0,0,0,0,0,0,0],[0,0,0,0,0,0,0,0,0,0],[0,0,0,0,0,0,0,0,0,0],
                    [0,0,0,0,0,0,0,0,0,0],[0,0,0,0,0,0,0,0,0,0],[0,0,0,0,0,0,0,0,0,0],[0,0,0,0,0,0,0,0,0,0],
                    [0,0,0,0,0,0,0,0,0,0],[0,0,0,0,0,0,0,0,0,0]]
    fs = pd.DataFrame(startingdata, index=['Madison Moonshiners (Brady)','Frisco Fireballs (Ethan W.)',
                                           'Margaronas (Tyler)','Colorado Gin Enthusiast (Joel)','Karolina Keg Stands (JC)',
                                           'Dysdunctional Frunks (Jordan)','McCaffeine Free Beverages (McKenna)',
                                           'Team rep292 (Ryan)','Pearl Pilsners (Alex)','Team ethanjflynn (Ethan F.)'],
                  columns=['1st Place Votes','2nd Place Votes','3rd Place Votes','4th Place Votes','5th Place Votes',
                           '6th Place Votes','7th Place Votes','8th Place Votes','9th Place Votes', '10th Place Votes'])
    fs.index.name = 'Team Name'
    fs.to_csv('.\\finalSheet.csv')   #Creating a CSV from dataframe
    
    
    r'''
    This automatically sets the NFL Week, commenting out for testing purposes, replacing with enter week input
    
    #Setting NFL week based on current date
    today = date.today()
    #Setting NFL Week 1 information
    week=1
    testdatestart = date(2022, 9, 14)   #Starts on Wednesdays, which is when power rankings are announced
    testdateend = date(2022, 9, 20)   #Ends on Tuesdays, the day before power rankings announcement
    #Looping through each NFL Week to check what week we are on (slightly adjusted from NFL schedule)
    while week <= 17:   #Only 17 weeks in fantasy football season
        if today >= testdatestart and today <= testdateend:
            break   #The correct date and week was found
        elif today >= date(2023, 1, 11):
            raise Exception('Out of Season')
            break   #Fantasy football season ends early January
        else:      
            week+=1 #If todays date was not within the test date range, add 1 week to NFL Week, and test times
            testdatestart = testdatestart + timedelta(weeks=1)
            testdateend = testdateend + timedelta(weeks=1)
    '''   
        
    week = input('Please enter the NFL Week: ')
    
    r'''
    This automatically pulls the correct file from the downloads folder, where the excel sheet will be sent to. 
    Commenting it out and replacing with choose file option for sake of reproducing on another machine
    
    #Opening Power Rankings (PR) Excel File
    week = str(week)   #Must be string to concatenate
    for (root,dirs,files) in os.walk(r"C:\Users\Jordan Ramsey\Downloads", topdown=True):   
        for name in files:  #Looking through each file in "Downloads" folder, where the power rankings are sent
            if "Week "+week+" Power" in name:   #Selecting the correct file based on week
                break_both = True   #Adding a flag to break from both loops
                break    
        if break_both:   #Breaking from second for loop
            break
    '''
    filename = askopenfilename()
    rs = pd.read_excel(filename)   #Opening correct power rankings file
   
    
    '''Individual polls come back in a string format, seperated by ";". Splitting each vote into its own cell'''
    responses = len(rs.loc[:,'Week '+week+' Power Rankings'])   #Number of responses
    i = 0   #Using this variable to track and limit loops and find row number
    while i <= responses-1:   #Subtracting by 1 due to 0-based indexing
        pr = rs.at[i,'Week '+week+' Power Rankings']   #Grabbing individual poll
        pr = pr.split(';')   #Spliting each response into respective cells
        pr.pop(10)   #Each poll ended with a ";", which created an extra list item. Popping out extra, empty item
        r = 0   #Tracking and limiting loops
        while r <= 9:   #Setting each name into respective cells. Each row is a new poll, each column is a new vote type
            ws.loc[i-1,r] = pr[r]
            r+=1
        i += 1


    '''Assigning point values to rankings and setting up tables for each vote level'''
    b=10   #This variable is the point multiplier for votes
    for a in range(10):   #Looping through each row on every column to find who got votes within each vote level (1st, 2nd, 3rd, etc.)
        add = ws[a].value_counts().rename_axis('team').reset_index(name=str(a)+'points')   #Counting how many times each name appears in each vote level
        count = len(add)   #Counting how many distinct teams in each vote level
        add[str(a)+'points'] *= b   #Multiplying each occurance of a vote by its specific vote level
        b-=1   #Each loop is a new vote level, decreasing vote multiplier
        names = ['Madison Moonshiners (Brady)','Frisco Fireballs (Ethan W.)','Margaronas (Tyler)',
                 'Colorado Gin Enthusiast (Joel)','Karolina Keg Stands (JC)','Dysdunctional Frunks (Jordan)',
                 'McCaffeine Free Beverages (McKenna)','Team rep292 (Ryan)','Pearl Pilsners (Alex)',
                 'Team ethanjflynn (Ethan F.)']   #Listing all league members to sort through with
        i=count   #Location marker to adjust which row to place data on for following loop
        for x in names:   #Looping through names list to check if team is already in the counted table
            if x in add.loc[:,'team']:   #if so, move onto next team and add to location marker
                i+=1
                continue
            else:   #if not, add team to table and add to location marker
                add.loc[i,'team']=x
                add.loc[i,str(a)+'points'] = 0
                i+=1
        add.to_csv('.\PR '+str(a)+r' Place Points.csv', index=False)   #Creating a new CSV for each vote level, that will be merged
    
        
    '''Merging Tables'''    
    b = 1   #Loop tracker and limiter
    for a in range(9):   #Each loop joins a vote level table with the preceeding vote level table (1st joins with 2nd, which joins with 3rd, etc)
        lfile = pd.read_csv('.\PR '+str(a)+' Place Points.csv')
        rfile = pd.read_csv('.\PR '+str(b)+' Place Points.csv')
        points = pd.merge(left = lfile, right = rfile, how = 'outer', left_on = 'team', right_on = 'team',
                          suffixes=('', '_drop')).filter(regex='^(?!.*_drop)')
        #Using an outer merge here due to how I format later on. Merge created duplicates and outer merge at least put the rows that I needed in predictible spots
        points.to_csv('.\PR '+str(b)+' Place Points.csv', index=False)   #Sending each new table back to CSV
        b+=1


    '''Removing Duplicates'''
    teams = list(points.iloc[:,0])   #Creating a list of all teams in table, including duplicates, to remove unnecesary rows
    teamsplus = list(points.iloc[:,0])   #Making a copy of list that will not have rows removed for the sake of iterating through while keeping the same index
    count = len(points)   #Number of total rows
    i = 1   #Starting at 1 because the first row of each new team is the row I need
    j=1   #Second tracker to ensure the rows I need will not be dropped
    for a in teamsplus:   #Iterating through list of team names in the order that they are in within the table. Using 2 lists in the loop because the list will re-index with each pop command
        occur = teams.count(a)   #Finding how many times a given name is duplicated in the table
        if i == count:   #Breaking loop when tracker gets past final row
            break
        elif occur > 1:   #If the team name is in the table more than once, this drops the 2nd occurance of that name (1st occurance is the one I need)
            points.drop(i, inplace=True, axis=0)
            teams.pop(j)   #Must use "j" here because list re-indexes with each pop command
            i+=1
        elif occur == 1:   #If the team name only occurs once, move on to the next name
            i+=1
            j+=1   #"j" will ensure it not only moves to the next team name, but also that the second occurance is the one that will be dropped

   
    '''We now have the correct point values assigned to the correct teams'''


    '''Formatting Tables'''
    fs = pd.read_csv(r'.\finalSheet.csv')
    points.sort_values(['team'], axis=0, ascending=[False], inplace=True, ignore_index=True)   #Ordering table based on team name   
    points.reset_index(drop=True,inplace=True)   #Replacing inaccurate index from removal of duplicates
    fs.sort_values(['Team Name'], axis=0, ascending=[False], inplace=True) #Doing the same with the final sheet so it matches with the points dataframe
    fs.reset_index(drop=True,inplace=True)


    '''Moving Calculated Data to Final Table'''
    data = points.iloc[0:10,1:12]   #Selecting points that the needed data is contained
    fs.iloc[0:10,1:12] = data   #Placing that data into the "fs" dataframe
    points.to_csv('.\PR 9 Place Points.csv', index=False)   #Finished with the points dataframe, sending back to CSV


    '''Adding a total points column and creating list & table for final rankings calculation'''
    i=0
    while i <= 9:   #Looping through each row to add the sum to new column
        fs.loc[i,12] = sum(fs.iloc[i,1:11])
        i+=1        
    fs.rename(columns={12:'Total Points'}, inplace=True)
    fs.sort_values(['Total Points'], axis=0, ascending=False, inplace=True)   #Putting table in ranking order
    rankingsordered = list(fs.iloc[:,0])   #Creating a list that is the team names
    rankingsordered.sort()   #Sorting that list alphabetically
    add = ws.iloc[:,0].value_counts().rename_axis('team').reset_index(name='votes')   #Creating table that shows 1st place votes per team
    add.sort_values(['team'], axis=0, ascending=True, inplace=True)   #Sorting that table alphabetically
 
    
    '''Giving teams who received 1st place votes an indicator of how many - will use later on final graphic'''
    z=0
    withvotes = list()   #Blank list to add onto
    for x in rankingsordered:   #Looping through each team name
        if x in list(add.loc[:,'team']):   #If their name is in the 1st place votes table
            y = str(add.loc[z,'team'])+' ('+str(add.loc[z,'votes'])+')'   #Add an indicator of how many votes they got
            withvotes.append(y)   #And add them to the list
            z+=1
        else:
            withvotes.append(x)   #If they did not receive a 1st place vote, add to list without indicator
     
    
    '''Ordering previously created list in ranking order & sending completed table to CSV'''
    withvotes.sort()   #Sorting alphabetically to match with total points
    fs.sort_values(['Team Name'], axis=0, ascending=[True], inplace=True)   #Sorting table alphabetically by team name to match list  
    fs.to_csv('.\\finalSheet.csv', index=False)   #Sending final rankings table to CSV to be used by rStudio to create graph
    pointsdata = pd.DataFrame(withvotes, columns=['team'])   #Combining team names including 1st place votes with total points
    pointsdata.loc[:,0] = list(fs.loc[:,'Total Points'])   #Adding total points (this was put in the same order as the team names earlier)
    pointsdata.sort_values([0], axis=0, ascending=[False], inplace=True)   #Putting in ranking order
    rankings = list(pointsdata.loc[:,'team'])


    '''Data is correctly formatted and can now by sent to rStudio for creating the stacked bar chart'''  
    r = robjects.r
    r.source('.\graph_creation.R')   #R is better for creating and formatting graphs/charts, all other code is contained within Python


    '''Opening the graph from rStudio along with the base graphic to place graph and rankings on'''
    graph = image.open('.\graph.png')
    graphic = image.open('.\graphic.png')


    '''Resizing & pasting graph onto graphic'''
    newsize = (1300,944)
    graph = graph.resize(newsize)
    graphic.paste(graph,(805,265))   #Exact placement for graph on graphic


    '''Adding Text to Image'''
    gtext = draw.Draw(graphic)
    titleFont = ImageFont.truetype(".\\URW Grotesk Regular.ttf", 108)   #Choosing font style & size for both title and rankings
    rankFont = ImageFont.truetype(".\\URW Grotesk Regular.ttf", 30)
    gtext.text((22,50), "Designated Drinker Week "+week+" Power Rankings", font=titleFont, fill=(255,255,255))   #Placing title on graphic
    x=116   #"x" and "y" are coordinates for first place team name on graphic
    y=273
    a=1
    for z in rankings:
        if a == 10:
            gtext.text((x+20,y), z, font=rankFont, fill=(255,255,255))
            a+=1
        else:
            gtext.text((x,y), z, font=rankFont, fill=(255,255,255))
            y+=98
            a+=1
            
    
    '''Displaying & Saving Image'''
    graphic.show()
    graphic.save(".\GraphicFinal.png")    
    graphic.save(r"C:\Users\Jordan Ramsey\iCloudDrive\Personal\Fantasy Football\Power Rankings\W"+week+" Graphic.png") #Need to change or delete this for the file to run correctly. This line sends the graphic to my phone

    
    '''For use next season, when full automation is set up'''
    #Calling browser script    
    #exec(open("C:\\Users\\Jordan Ramsey\\iCloudDrive\\Personal\\Fantasy Football\\Python Excel Files\\AutoBrowser.py").read())


Main()





























