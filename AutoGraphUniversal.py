#Loading Packages
import pandas as pd
from datetime import date, timedelta
import os
import rpy2.robjects as robjects
from PIL import Image as image
from PIL import ImageDraw as draw
from PIL import ImageFont
from rpy2.robjects.packages import importr
from tkinter.filedialog import askopenfilename

    
'''Setting NFL week based on current date'''
def set_week():
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
    return str(week)


'''Creating Working & Final CSV & MS Forms Power Rankings Excel File'''
def create_csv(week):
    #Creating working sheet (ws) to perform calculations and data transformation on and creating CSV
    ws = pd.DataFrame()
    ws.to_csv('workingSheet.csv')
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
    fs.to_csv('finalSheet.csv')   #Creating a CSV from dataframe
    fs = pd.read_csv('finalSheet.csv')

    filename = askopenfilename()
    rs = pd.read_excel(filename)
    return ws, fs, rs
   

'''Individual polls come back in a string format, seperated by ";". Splitting each vote into its own cell'''
def split(rs):
    responses = len(rs.loc[:,f'Week {week} Power Rankings'])   #Number of responses
    for i in range(responses):   #Subtracting by 1 due to 0-based indexing
        pr = rs.at[i,f'Week {week} Power Rankings']   #Grabbing individual poll
        pr = pr.split(';')   #Spliting each response into respective cells
        pr.pop(10)   #Each poll ended with a ";", which created an extra list item. Popping out extra, empty item
        r = 0   #Tracking and limiting loops
        while r <= 9:   #Setting each name into respective cells. Each row is a new poll, each column is a new vote type
            ws.loc[i-1,r] = pr[r]
            r+=1
        i += 1


'''Assigning point values to rankings and setting up tables for each vote level'''
def assign_points():
    b=10   #This variable is the point multiplier for votes
    for a in range(10):   #Looping through each row on every column to find who got votes within each vote level (1st, 2nd, 3rd, etc.)
        add = ws[a].value_counts().rename_axis('team').reset_index(name=f'{str(a)}points')   #Counting how many times each name appears in each vote level
        count = len(add)   #Counting how many distinct teams in each vote level
        add[f'{str(a)}points'] *= b   #Multiplying each occurance of a vote by its specific vote level
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
                add.loc[i,f'{str(a)}points'] = 0
                i+=1
        add.to_csv(f'PR {str(a)} Place Points.csv', index=False)   #Creating a new CSV for each vote level, that will be merged

    
'''Merging Tables'''  
def merge():  
    b = 1   #Loop tracker and limiter
    for a in range(9):   #Each loop joins a vote level table with the preceeding vote level table (1st joins with 2nd, which joins with 3rd, etc)
        lfile = pd.read_csv(f'PR {str(a)} Place Points.csv')
        rfile = pd.read_csv(f'PR {str(b)} Place Points.csv')
        points = pd.merge(left = lfile, right = rfile, how = 'outer', left_on = 'team', right_on = 'team',
                          suffixes=('', '_drop')).filter(regex='^(?!.*_drop)')
        #Using an outer merge here due to how I format later on. Merge created duplicates and outer merge at least put the rows that I needed in predictible spots
        points.to_csv(f'PR {str(b)} Place Points.csv', index=False)   #Sending each new table back to CSV
        b+=1
        os.remove(f'PR {str(a)} Place Points.csv')
    os.remove('PR 9 Place Points.csv')
    return points


'''Removing Duplicates'''
def rem_dup(points):
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
    return teams, teamsplus

   
'''We now have the correct point values assigned to the correct teams'''


'''Formatting Tables'''
def formatting():
    fs = pd.read_csv('finalSheet.csv')
    points.sort_values(['team'], axis=0, ascending=[False], inplace=True, ignore_index=True)   #Ordering table based on team name   
    points.reset_index(drop=True,inplace=True)   #Replacing inaccurate index from removal of duplicates
    fs.sort_values(['Team Name'], axis=0, ascending=[False], inplace=True) #Doing the same with the final sheet so it matches with the points dataframe
    fs.reset_index(drop=True,inplace=True)
    return points, fs


'''Moving Calculated Data to Final Table'''
def moveto_final(points, fs):
    data = points.iloc[0:10,1:12]   #Selecting points that the needed data is contained
    fs.iloc[0:10,1:12] = data   #Placing that data into the "fs" dataframe
    print(points)
    print(fs)


'''Adding a total points column and creating list & table for final rankings calculation'''
def total_points():
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
    return add, rankingsordered
 

'''Giving teams who received 1st place votes an indicator of how many - will use later on final graphic'''
def firstpl_ind(add, rankingsordered):
    z=0
    withvotes = list()   #Blank list to add onto
    for x in rankingsordered:   #Looping through each team name
        if x in list(add.loc[:,'team']):   #If their name is in the 1st place votes table
            y = str(add.loc[z,'team'])+' ('+str(add.loc[z,'votes'])+')'   #Add an indicator of how many votes they got
            withvotes.append(y)   #And add them to the list
            z+=1
        else:
            withvotes.append(x)   #If they did not receive a 1st place vote, add to list without indicator
    return withvotes
 

'''Creating ordered list in ranking order & sending completed table to CSV'''
def order_tocsv(withvotes):
    withvotes.sort()   #Sorting alphabetically to match with total points
    fs.sort_values(['Team Name'], axis=0, ascending=[True], inplace=True)   #Sorting table alphabetically by team name to match list  
    fs.to_csv('finalSheet.csv', index=False)   #Sending final rankings table to CSV to be used by rStudio to create graph
    pointsdata = pd.DataFrame(withvotes, columns=['team'])   #Combining team names including 1st place votes with total points
    pointsdata.loc[:,0] = list(fs.loc[:,'Total Points'])   #Adding total points (this was put in the same order as the team names earlier)
    pointsdata.sort_values([0], axis=0, ascending=[False], inplace=True)   #Putting in ranking order
    rankings = list(pointsdata.loc[:,'team'])
    return rankings

'''Data is correctly formatted and can be placed into a stacked bar chart using R'''
def r_graph():
    importr("ggplot2")
    importr('tidyr')
    importr('RColorBrewer')
    importr('ggrepel')


    robjects.r(r'''
            #Setting working directory. To have this work on a different machine, switch this to that machines marking directory filepath
            setwd("C:\\Users\\Jordan Ramsey\\Documents\\PR Python CSVs")
            getwd()
            
            #Opening Files
            points = read.csv(file = "finalSheet.csv", sep=',')
            
            #Renaming columns for proper indexing
            i=1
            x=2
            while (i <=9)
            {
              colnames(points)[x] = paste('X0',i,'_Place_Votes', sep='')
              i=i+1
              x=x+1
            }
            colnames(points)[11] = 'X10_Place_Votes'
            
            #Creating Plot
            pointslong = pivot_longer(points, X01_Place_Votes:X10_Place_Votes, names_to='VoteType', values_to='points')
            graph = ggplot(data=pointslong, aes(x=reorder(Team.Name, -Total.Points), y=points, fill=VoteType)) +
              geom_bar(position='stack',stat='identity',width=.5)
            
            #Formatting Plot
            graph = graph + labs(x='Team Name') + scale_fill_discrete(name='Vote Type', labels=c('1st Place Votes','2nd Place Votes','3rd Place Votes','4th Place Votes','5th Place Votes','6th Place Votes','7th Place Votes','8th Place Votes','9th Place Votes','10th Place Votes'))
            graph = graph + theme(panel.background = element_rect(fill='#353332'), plot.background = element_rect(fill='#353332'))
            graph = graph + theme(panel.grid.major = element_blank(), panel.grid.minor.y = element_blank())
            graph = graph + theme(legend.background = element_rect(fill='#353332'), legend.key = element_rect(fill='#353332'))
            graph = graph + theme(legend.position = 'top')
            graph = graph + theme(axis.text.x = element_text(size=10,color='white'), axis.title.x = element_text(size=20,color='white'),axis.text.y = element_text(size=20,color='white'), axis.title.y = element_text(size=20,color='white'), legend.text = element_text(color='white'))
            graph = graph + theme(legend.key.size = unit(2, 'cm'),
                                  legend.key.height = unit(2, 'cm'),
                                  legend.key.width = unit(2, 'cm'),
                                  legend.text = element_text(size=20),
                                  legend.title = element_text(size=20,color='white'))
            
            #Saving plot and exporting to computer
            png(file="graph.png", width=1712, height=1001)
            print(graph)
            dev.off()
    ''')


'''Resizing & pasting graph onto graphic & Saving'''
def graphic(rankings):
    graph = image.open('graph.png')
    graphic = image.open('graphic.png')
    newsize = (1300,944)
    graph = graph.resize(newsize)
    graphic.paste(graph,(805,265))   #Exact placement for graph on graphic
    '''Adding Text to Image'''
    gtext = draw.Draw(graphic)
    titleFont = ImageFont.truetype("URW Grotesk Regular.ttf", 108)   #Choosing font style & size for both title and rankings
    rankFont = ImageFont.truetype("URW Grotesk Regular.ttf", 30)
    gtext.text((22,50), f"Designated Drinker Week {week} Power Rankings", font=titleFont, fill=(255,255,255))   #Placing title on graphic
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
    graphic.show()   #Displaying Image
    graphic.save("GraphicFinal.png")   #Saving Image 
    graphic.save(f"C:\\Users\\Jordan Ramsey\\iCloudDrive\\Personal\\Fantasy Football\\Power Rankings\\W{week} Graphic.png") #Need to change or delete this for the file to run correctly. This line sends the graphic to my phone




#week = set_week()   #Comment out for testing
week = input('Enter NFL Week: ')   #Uncomment for testing

ws, fs, rs = create_csv(week)

split(rs)

assign_points()

points = merge()

teams, teamsplus = rem_dup(points)

points, fs = formatting()

moveto_final(points, fs)

add, rankingsordered = total_points()

withvotes = firstpl_ind(add, rankingsordered)

rankings = order_tocsv(withvotes)

r_graph()

graphic(rankings)

'''For use next season, when full automation is set up'''
#Calling browser script    
#exec(open("C:\\Users\\Jordan Ramsey\\iCloudDrive\\Personal\\Fantasy Football\\Python Excel Files\\AutoBrowser.py").read())




























