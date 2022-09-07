# The Weakest Link - Built In Excel For Zoom Quiz

## Background
At the start of covid as everyone was doing zoom quizzes I had the idea of creating The Weakest Link in Excel for something a bit different.

The Weakest Link is a British quiz show which used to be hosted by Anne Robinson during my childhood: https://en.wikipedia.org/wiki/Weakest_Link

This took me about a weekend to build and was a hit during the 3/4 zoom quizzes I did with various friends and family I did the quiz with.

While the boom in zoom quizzes has long gone I thought this would be good to share with the community better late than never.

## Some Notes
1. Why did I build it in Excel? I wanted to build something quickly and while the web would have been perfect to share with others, my job at the time was a professional Excel developer so this technology had the quickest "time to market"
2. What versions of Excel does this work with? I built this in Excel 2016 but it should work in any version after 2013 (this hasn't been tested)
3. Can I look through the code? Yes - I've left the project fully unlocked so feel free to have a look and edit anything

## Getting Started And Running Your Own Quiz

1. Download the Excel file WeakestLink.xlsb
1. This is ideally run on a computer with two screens. This will allow the quiz master to have one window looking at the Quiz Master Dashboard on the Ref sheet and the another window on the Main sheet for the players. Share the window showing the Main screen on zoom with all the players. To create a new window in Excel press View > New window
2. On the Main sheet, enter the name of the contestants in the row which says "Enter names in this row:". This game allows up to 8 players
3. Set the length of each round (in seconds) in the named range "RoundLength" on the Ref sheet
4. Press the button new game. This will clear the previous game's state, ie things like previous questions asked, who was eliminated etc
5. To start a new round press the "Next Round" button. This will start the countdown timer.
6. Use the green/red/bank buttons to mark when a contestant gets a question right/wrong/wants to bank
7. At the end of the round a message box will pop up saying it is the end of the round
8. When it is time to vote for who the weakest link is, you can use the "Player X" buttons under each player to keep track of the number of votes a player has received
9. To eliminate the player with the most votes, press the "Say Goodbye" button. If there is a tie for the most votes, the strongest link has the deciding vote. In the input box that pops up, enter the player number of the player eliminated


## Adding New Questions
1. Extend the named range Questions if it doesn't have enough space for your questions.
2. Add your questions and answers in this range
3. Make sure the id column runs sequentially from 1 to [ the number of questions you have ]
4. Double check the value in the named range "NumberOfQuestions" matches the number of questions you have
5. If you want to a round to only have a specific range of questions, you can use the table in the named range "QuestionIdLimits" to set this.
Eg if Round 1 has limits 1 and 20, in Round 1, questions will only be generated between ids 1 and 20.

