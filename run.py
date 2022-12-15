from create_df import setup_rankings, assign_points
from r_graph import create_graph
from create_graphic import first_ind, set_week, graphic


def main():
    '''Calling all functions from other files that makeup this project.'''
    team_names, polls, num_teams = setup_rankings()    
    assign_points(team_names, polls, num_teams)    
    create_graph()    
    rankings = first_ind()    
    week = set_week()
    
    # Currently, creating the graphic only supports 10 teams.
    # As developments are made to this script, support for 6, 8, 10, and 12
    # member leagues will be implemented and their functions called accordingly
    if len(rankings) == 10:
        graphic(rankings, week)
    else:
        raise Exception('Not 10 Teams')

    
if __name__ == '__main__':
    main()