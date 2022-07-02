# Stock Analysis: Refactored Code
## Overview of the project
### Purpose and Background
To help Steve analyze and find the best stocks for his parents to invest in, I developed a VBA code via Excel that would analyze the data efficiently. Since Steve was so content with the initial VBA code he asked me to improve it so he could analyze the entire stock market over the last few years. While the initial code may have worked, there was room for improvement so that it would run faster and more efficiently. In this project, I refactored the code to make it run faster. I have provided images of my improved code and timestamps of execution times to compare the original code with the refactored verison. 
## Results
### Analysis 
The goal for this project was to create a new code for Steve so that he could analyze all stocks throughout multiple years. Since I had already developed a code for him, all I had to do was slightly modify it so that not only could it do what Steve wanted it to, but in a much more efficient manner. Some modifications in my refatored code can be seen in For Loop funncitons of the tickerVolumes and the tickerIndex. Also, in this refactored code, I left the For Loop in "i" and removed the "j" variable. One key component in the refactor is that I also reassigned the buttons on the worksheet to operate under the refactored code and not the previous code. If not, it would keep giving me the data under the original code. All of this led to the code running faster and efficiently. I have attached screenshots of time stamps from the original code and the refactored for comparison. The images indicate that the refactored code indeed ran faster than the original. 

This is a copy of my refactored code script. 

<img width="683" alt="Screen Shot 2022-07-01 at 5 21 27 PM" src="https://user-images.githubusercontent.com/107595127/176979987-197943c4-3e8e-4721-b915-0c3c658d78ae.png">

These are screenshots of the time execution from the original code.

<img width="257" alt="Screen Shot 2022-07-01 at 2 45 18 PM" src="https://user-images.githubusercontent.com/107595127/176980479-76f153b9-72ce-418c-bbe5-592be38da87a.png">
<img width="260" alt="Screen Shot 2022-07-01 at 2 45 28 PM" src="https://user-images.githubusercontent.com/107595127/176980480-9c0031fc-dd8d-4649-ba7f-7674f7498023.png">

These are screenshots of the time execution from the refactored verison of the code. As you can see, the refactored code ran faster. 

<img width="252" alt="Screen Shot 2022-07-01 at 4 55 52 PM" src="https://user-images.githubusercontent.com/107595127/176980526-06abcf48-2c3f-45d5-b0cd-2418f7197ff0.png">
<img width="255" alt="Screen Shot 2022-07-01 at 4 56 06 PM" src="https://user-images.githubusercontent.com/107595127/176980528-4453566c-644f-44d3-bfa1-067dce69bc90.png">

## Summary
### Advantages and Disadvantages of Refactoring Code
As the analysis and images show, refactoring a code can improve the efficiency and time in which a code runs. It can allow for a cleaner layout so that others can easily follow or read your code. In doing so, if you (the developer) ever needed to return to the code and re-modify it, you won't get as lost. Overall, it improves one's ability to read, debug, and for faster programming. Some cons can include losing infomration that you once had when it was writted out step by step. But at this point, we'd expect one to know exactly what they are doing. 
### Advantages and Disadvantages of Refactored Stock Analysis
As previously mentioned, one of the greatest advantage of refactoring this particular code is that the refactored code made the program run significantly faster. It also looks a lot cleaner and simplified. Also, we are not able to analyze the entire stock market data over multiple years for Steve. A disadvantage is that all the simplification may make it seem as if there are key components that are missing, but if done correctly, they will still exist within the code.
