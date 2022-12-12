install.packages("ggplot2")
library(ggplot2)
install.packages('tidyr')
library(tidyr)
install.packages('RColorBrewer')
library(RColorBrewer)
install.packages('ggrepel')
library(ggrepel)

#Setting working directory. To have this work on a different machine, switch this to that machines marking directory filepath
setwd("C:\\Users\\Jordan Ramsey\\Documents\\PR Project Files")
getwd()

#Opening Files
points = read.csv(file = "./finalSheet.csv", sep=',')

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
png(file="./graph.png", width=1712, height=1001)
print(graph)
dev.off()
