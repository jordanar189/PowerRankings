import pandas as pd
from datetime import date, timedelta
from PIL import Image as image
from PIL import ImageDraw as draw
from PIL import ImageFont



def first_ind():
    '''Giving teams who received 1st place votes an indicator of
     how many - will use soon on final graphic.
     '''
    rankings = pd.read_csv("Tables/points_calculated.csv")
    rankings.drop('Unnamed: 0', axis=1, inplace=True)   # Removing added index
    teams = rankings.loc[:,'Team Name']
    teams = list(teams)
    # Dividing 1st place points by number of teams to find first
    # place votes per team
    for name in teams:   
        if int(rankings.loc[rankings['Team Name'] == name,
                            '01 Place Points']) > 0:
            votes = int(rankings.loc[rankings['Team Name'] == name,
                                     '01 Place Points']) / len(teams)
            votes = int(votes)
            rankings.loc[rankings['Team Name'] == name,
                         'Team Name'] = f'{name} ({votes})'
    ind_rankings = rankings.loc[:,'Team Name']
    ind_rankings = list(ind_rankings)   # List with 1st place indicator
    return ind_rankings

def set_week():
    '''Using datetime to set NFL week based on current date, to be used in
    the title of the graphic.
    '''
    today = date.today()
    week = 1
    testdatestart = date(2022, 9, 14)
    testdateend = date(2022, 9, 20)
    # Checking each week, breaking when today's date falls within
    # the week being tested
    while week <= 17:   
        if today >= testdatestart and today <= testdateend:
            break
        elif today >= date(2023, 1, 11):   # Sends error if out of season
            raise Exception('Out of Season')
            break
        else:      
            week+=1
            testdatestart = testdatestart + timedelta(weeks=1)
            testdateend = testdateend + timedelta(weeks=1)
    return str(week)
    

def graphic(rankings, week):
    '''Taking all collected content (totaled points, first place votes,
    stacked bar chart) and setting onto graphic background.
    '''
    graph = image.open('Pictures/graph.png')
    graphic = image.open('Pictures/graphic.png')
    newsize = (1300,944)
    graph = graph.resize(newsize)
    graphic.paste(graph,(805,265))  # Placing graph on graphic
    
    # Placing title onto graphic
    add_text = draw.Draw(graphic)
    titleFont = ImageFont.truetype("Fonts/URW Grotesk Regular.ttf", 108)
    rankFont = ImageFont.truetype("Fonts/URW Grotesk Regular.ttf", 30)
    add_text.text((22,50), f"Designated Drinker Week {week} Power Rankings", 
               font=titleFont, fill=(255,255,255))
    
    # Placing text onto graphic in order of rankings
    x=116
    y=273
    a=1
    for z in rankings:
        if a == 10:
            add_text.text((x+20,y), z, font=rankFont, fill=(255,255,255))
            a+=1
        else:
            add_text.text((x,y), z, font=rankFont, fill=(255,255,255))
            y+=98
            a+=1
    graphic.show()
    graphic.save("Pictures/GraphicFinal.png")
    graphic.save("C:\\Users\\Jordan Ramsey\\iCloudDrive\\Personal\\" +
                 f"Fantasy Football\\Power Rankings\\W{week} Graphic.png")