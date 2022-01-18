# Module 2 Challenge: VBA Refactoring Project/Challenge Analysis:

###### An analysis of the data for module 2 challenge.

## Overview of the Project

### Purpose:

In module 2, I used VBA to automate excel formulas in order to analyze stock info. Although the 2 versions of analysis that I conducted (original module analysis and the VBA_Challenge analysis) gave me the same output:


![image](https://user-images.githubusercontent.com/96212660/149862598-893e7821-24ee-4930-9ea1-f6c6bd43e480.png)


![image](https://user-images.githubusercontent.com/96212660/149862634-c195ab9e-0ce2-4aac-95de-74357b7af006.png)


I wanted to refactor the code to loop through all the data provided in the [VBA Challenge Data Workbook](https://github.com/ayyray/stock-analysis/blob/03995b9e922d6c6183b520cf6c85b52aac64716f/VBA_Challenge%20.xlsm). Through the process of refactoring my code, my intention was to see if the refactored code was any faster (the VBA code that was scripted to create macros for my analysis). 

### Background:

The code that I spent time putting together for module 2 ran an analysis of the provided stock information. I wanted to analyze a handful of stocks using excel. VBA was used to help make automated analysis so that I could reuse it (loop) through the data in order to ensure there was logical and analytical consistency. The refactoring served the purpose of improved efficiency for me writing it and for improved functionality (time stamp outcome or for a third party to reference.) 


## Results:

### Analysis of Results:

Results of my findings for my initial code (green_stocks.xlm) for 2017 and 2018 are as followed:

![green_stocks_2017](https://user-images.githubusercontent.com/96212660/149863847-8c649f13-cf6e-4ff9-8550-90dbb8f40d97.png)


![green_stocks_2018](https://user-images.githubusercontent.com/96212660/149863872-798a744e-e40b-495c-b453-bfb718e9720e.png)


Results of my finding for refactored code (VBA_Challenge) 2017 and 2018 are as followed:

<img width="246" alt="VBA_Challenge_2017 " src="https://user-images.githubusercontent.com/96212660/149864013-714e6ed0-168f-4634-8f2c-a1cdb5a6003e.png">


<img width="247" alt="VBA_Challenge_2018 " src="https://user-images.githubusercontent.com/96212660/149864031-83d10bae-f200-4a18-8089-000faca5bf30.png">


First glance of the outputs show that the refactored code was in fact faster for both 2017 and 2018 in the VBA_Challenge refactored code. The biggest part of refactoring did not come from rewriting the code or even moving around the formatting of the code. Refactoring in this analysis came down to the utilization of well placed for loops. 

-------------------------------------------------------------------------------------------------------------------------------------------------------------------

#### Here are some of the following pieces of the code that I believe made for a more efficient and quicker output for the 2 scripts:

##### In my initial code when I wanted to find the total Volume for a ticker, I used the following code:

![image](https://user-images.githubusercontent.com/96212660/149864124-dd2d137f-16b2-4be0-962e-4dcce4698a3d.png)

##### Refactored code:

![image](https://user-images.githubusercontent.com/96212660/149864142-b5863a02-e5fa-41bc-a9fb-509cb64455a0.png)


#### Analysis of Refactored Code: 
While I used the same process of utilizing for loops here, I think the refactored code made it easier to visualize what analysis was going to happen and it also cleaned up the code. By creating a ticker Index and setting it equal to zero to start, I was then able to create a for loop for the ticker Volumes and then use another loop to increase the volume for the ticker. (I will show the refactored increase volume next.) As opposed to the original script where I had to first create a loop through the tickers, set the ticker equal to the ticker(i), and set volume = 0, I was able to cut out a step.

-------------------------------------------------------------------------------------------------------------------------------------------------------------------




##### In my initial code when I wanted to determine volume, starting, and ending price for tickers I used the following code:

![image](https://user-images.githubusercontent.com/96212660/149864287-798eee38-f89a-4eed-8936-152f9a4a70d7.png)
 
##### Refactored code:

![image](https://user-images.githubusercontent.com/96212660/149864310-dc3799fb-62c0-462c-bfac-6cc57ec269b3.png)


#### Analysis of Refactored Code: 
For this section the code is similar with the difference of including the index to loop over the stocks. I found this helpful when I wanted to go back and reference the data to see the logic behind my analysis. I found this helpful in that if I wanted to go over a larger dataset (example stocks), I could adjust the value to where I was defining the index and expand it. It also made it easier for me to follow along with the code so I was not getting lost. It can easily begin to become a blob of numbers/data, but having the index labeled in my formulas allows me to just go up and reference what my for loop will be going through.

-------------------------------------------------------------------------------------------------------------------------------------------------------------------

##### The original code in which I reafactored the outputs formulas:

![image](https://user-images.githubusercontent.com/96212660/149864423-d21d464f-1ad9-45f4-90fb-4f64588b742c.png)

##### Refactored code:

![image](https://user-images.githubusercontent.com/96212660/149864468-b6373295-1166-4c44-90d4-c5f5fcf29049.png)
 

#### Analysis of Refactored Code: 
Looping through the arrays (that were defined earlier) were as simple as adding an index to the arrays. I found this to be helpful in once again assisting in clarity for reference. I can see that my index is defined to loop through 0 to 11. If I wanted to expand my set and loop through a larger data set, I could simply redefine the loop to account for those.

-------------------------------------------------------------------------------------------------------------------------------------------------------------------

### Summary:
After going through the process of writing a code and then going back and refactoring has both its advantages and disadvantages.

#### Advantages:
- Refactoring can help improve the code to make it more efficient and concise in terms of output. (As seen in Module Challenge). 
- It can help you gain a new perspective as a coder for improvements in your future as a coder. I always believe having a different perspective can prove to be extremely beneficial in some capacity.
- Refactoring can make your code or whatever program you might be creating run faster! Who doesn’t love efficiency!

#### Disadvantages:
- If I was writing code or a program and someone else came in and refactored it might mess up my train of thought and throw off my process.
- If you refactor and you are not exact you could risk affecting the entire program you are writing.
-I imagine spending so much time refactoring could eventually get to the point where its not benefiting the project and you are spending too much time fixated on a small part of a process that might shave off just a second of your overall output.

In terms of the refactoring that took place in my VBA analysis I could see more advantages than disadvantages. The type of coding that I was doing was for stock analysis. Although I am new to the realm of data analytics and VBA I know the projects I could see myself doing in the future will definitely entail much larger data sets than just 12 stocks. I would rather find efficient ways to refactor my code now to get the kinks out. That way if I run into a larger data set, I will have already exposed myself to some shortcuts/efficient approaches to handling the data.	








