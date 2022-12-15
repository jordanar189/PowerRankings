import pandas as pd


def setup_rankings():
    '''This function will open the MS Forms Power Rankings file compiled by
    Power Automate. The dataframe for totaling points is set up as well, with
    team names and placeholder data (0) put in.
    '''
    pr_sheet = pd.read_excel("Tables/Power Rankings.xlsx")
    polls = list(pr_sheet.loc[:,'Power Rankings'])
    #Individual polls come as strings, with ", " as seperator
    team_names = polls[0].split(', ')   
    num_teams = len(team_names)
    
    # Setting up working dataframe
    points_sheet = pd.DataFrame(index=range(num_teams))
    placeholder_data = list()
    while len(placeholder_data) < num_teams:   # Setting placeholder data
        placeholder_data.append(0)
    for x in range(num_teams):   #Placing in placeholder data
        if x < 9:
            points_sheet.insert(x, f'0{x+1} Place Points', placeholder_data)
        else:
            points_sheet.insert(x, f'{x+1} Place Points', placeholder_data)
    points_sheet.insert(0, 'Team Name', team_names)
    return points_sheet, polls, num_teams


def assign_points(points_sheet, polls, num_teams):
    '''To assign point to each team, this function will ilterate through
    each team name in each poll and assign a point value. Because each poll
    is provided in ranking order, the first name in each poll gets 10 points,
    the next 9 points, and so on.
    '''
    for ind_poll in polls:
        split_poll = ind_poll.split(', ')
        i = 1
        points = num_teams   #1st place point value is equal to number of teams
        for name in split_poll:
            if i < 10:
                points_sheet.loc[points_sheet['Team Name'] == name,
                                 f'0{i} Place Points'] += points
                i += 1
                points -= 1
            else:   #Only places 1-9 need the 0 in front
                points_sheet.loc[points_sheet['Team Name'] == name,
                                 f'{i} Place Points'] += points
                i += 1
                points -= 1
                
    # Adding total points column
    for x in range(num_teams):
        points_sheet.loc[x, 'Total Points'] = sum(points_sheet.iloc[x, 1:num_teams+1])
    points_sheet.sort_values(['Total Points'], axis=0, ascending=False,
                             inplace=True)
    
    # Saving to folder for R program to pick up
    points_sheet.to_csv('Tables/points_calculated.csv')
    return points_sheet