# Challenge2
## stock-analysis
## Overview of Project
Steve had loved the workbook that we created for him. Now he wants us to do a little more research for his parents which can expand the dataset to include the entire stock market over the last few years.

### Purpose
Since Steve is looking for growth opportunities outside of the green stock selections he was asking if we could update the VBA script to be used on all stocks on the market. The goal of this report was to analyze every stock in the market as efficiently as possible. We take 12 stocks and get the informations very quickly using the script we create.

### Results
Once we were provided the inital report we then could take that report and work on it in order to create a better more efficient one. Once we found we could run the script quicker by creating four new arrays. tickerIndex, tickerVolumes, tickerStartingPrice, and tickerEndPrice.

With this method it would pull the ticker symbol, starting price, ending, price, and volumes once before it goes through the remainder of the script. The method we used before would go through the whole script which would be more time consuming.

### Refactored Code Examples

![image](C:\Users\Zack.Computer\Desktop\Challenge2\Resources\VBA Script Refractoring.png)

**Refactored Script for Looping Through Iterations**

![image](C:\Users\Zack.Computer\Desktop\Challenge2\Resources\Refractored Script for Looping.png)

**Original Script for Looping Through Iterations**

![image](C:\Users\Zack.Computer\Desktop\Challenge2\Resources\Original Script for Looping.png)

The difference here is that the ticker from the original script was a variable. With the refractored code tickerIndex in referenced with each array, such as totalVolume(tickerIndex) and tickerStartingPrice(tickerIndex) which provided much quicker results.

### Speed of Macro Script

The original script ran to over 1 second for both years. When we refractored the code it made the script run at less than .2 seconds which would be roughly 80% faster.

**2017 Refactored Results**

![image](C:\Users\Zack.Computer\Desktop\Challenge2\Resources\VBA_Challenge_2017.png.png)

**2018 Refactored Results**

![image](C:\Users\Zack.Computer\Desktop\Challenge2\Resources\VBA_Challenge_2018.png.png)

## Stock Market Analysis Summary
    1. What are the advantages and disadvantages of refactoring code?
A major advantage of refactoring code is you do not need to start over everytime. It saves the programmer time and effort if they are starting a new project. A disadvantage would come when it is a code you did not write and you would not know if it would support the data you are working on.

    2. How do these advantages and disadvantages apply to the original VBA script?
Most of the work was already provided from this script. All we were asked to do was to make it more efficient. But having that code prewritten I did have to go through to make sure it would support the dataset we were given.