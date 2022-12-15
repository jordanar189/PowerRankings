import os
#Setting environment to use R with
os.environ['R_HOME'] = r'C:/Program Files/R/R-4.2.1'   
import rpy2.robjects as robjects
from rpy2.robjects.packages import importr


def create_graph():
    '''In order to make the stacked bar chart, this function will be using the
    programming language R. R can be the better option for data visualization
    and can also allow for easier, more direct coding to create a plot.
    
    This function will create the plot, do a bit of formatting to column names
    and then do a number of formatting to the plot to allow it to fit with the
    final graphic.
    ''' 
    # Importing R Packages
    importr('ggplot2')
    importr('tidyr')
    importr('RColorBrewer')
    importr('ggrepel')
    importr('tidyverse')
       
    robjects.r(r'''
               
        # Opening File
        points = read.csv(file='Tables/points_calculated.csv', sep=',')
        num_teams = length(points)
        
        # Creating Base Plot
        pointslong = pivot_longer(points, 4:num_teams-1, names_to='VoteType',
                                  values_to='points')
        graph = ggplot(data=pointslong, aes(x=reorder(Team.Name, -Total.Points),
                                            y=points, fill=VoteType)) +
          geom_bar(position='stack',stat='identity',width=.5) +
          scale_x_discrete(labels=function(x) str_wrap(x, width = 10))
        
        # Removing X from column labels that gets added by R
        col_lbls = colnames(points[0,4:num_teams-1])
        col_lbls = list(col_lbls)
        for (x in col_lbls) {
          x = sub('X', '', x)
        }
                
        # Formatting Plot
        graph = graph + labs(x='Team Name') +
        scale_fill_discrete(name='Vote Type', labels=x) +
        theme(panel.background = element_rect(fill='#353332'),
              plot.background = element_rect(fill='#353332'),
              panel.grid.major = element_blank(),
              panel.grid.minor.y = element_blank(),
              legend.background = element_rect(fill='#353332'),
              legend.key = element_rect(fill='#353332'),
              legend.position = 'top',
              axis.text.x = element_text(size=20,color='white'),
              axis.title.x = element_blank(),
              axis.text.y = element_text(size=20,color='white'),
              axis.title.y = element_blank(),
              legend.text = element_text(color='white', size=15),
              legend.key.size = unit(1.5, 'cm'),
              legend.key.height = unit(1.5, 'cm'),
              legend.key.width = unit(1.5, 'cm'),
              legend.title = element_text(size=15,color='white'))
        
        #Saving plot and exporting to computer
        png(file='Pictures/graph.png', width=1300, height=944)
        print(graph)
        dev.off()
        ''')