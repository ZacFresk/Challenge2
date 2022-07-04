## stock-analysis
## Overview of Project
Steve had loved the workbook that we created for him. Now he wants us to do a little more research for his parents which can expand the dataset to include the entire stock market over the last few years.

### Purpose
Since Steve is looking for growth opportunities outside of the green stock selections he was asking if we could update the VBA script to be used on all stocks on the market. The goal of this report was to analyze every stock in the market as efficiently as possible. We take 12 stocks and get the informations very quickly using the script we create.

### Results
Once we were provided the inital report we then could take that report and work on it in order to create a better more efficient one. Once we found we could run the script quicker by creating four new arrays. tickerIndex, tickerVolumes, tickerStartingPrice, and tickerEndPrice.

With this method it would pull the ticker symbol, starting price, ending, price, and volumes once before it goes through the remainder of the script. The method we used before would go through the whole script which would be more time consuming.

### Refactored Code Examples

![VBA Script Refractoring](https://user-images.githubusercontent.com/107363203/177089573-8b989519-9fba-4bac-8a07-b71da08f04d4.png)

**Refactored Script for Looping Through Iterations**

![Refractored Script for Looping](https://user-images.githubusercontent.com/107363203/177089600-7a3bddcb-52da-4ca9-9d5b-d47cd782ded2.png)


**Original Script for Looping Through Iterations**

![Original Script for Looping](https://user-images.githubusercontent.com/107363203/177089611-01496b1d-35bc-4714-85d5-147b6f25df0c.png)

The difference here is that the ticker from the original script was a variable. With the refractored code tickerIndex in referenced with each array, such as totalVolume(tickerIndex) and tickerStartingPrice(tickerIndex) which provided much quicker results.

### Speed of Macro Script

The original script ran to over 1 second for both years. When we refractored the code it made the script run at less than .2 seconds which would be roughly 80% faster.

**2017 Refactored Results**

![VBA_Challenge_2017 png](https://user-images.githubusercontent.com/107363203/177089636-f17bd0a7-5a3f-4b8d-8c71-325c6fef6e97.png)

**2018 Refactored Results**

![VBA_Challenge_2018 png](https://user-images.githubusercontent.com/107363203/177089655-f652f1a6-6b47-44a0-8aa3-170c522cb569.png)

## Stock Market Analysis Summary
    1. What are the advantages and disadvantages of refactoring code?
A major advantage of refactoring code is you do not need to start over everytime. It saves the programmer time and effort if they are starting a new project. A disadvantage would come when it is a code you did not write and you would not know if it would support the data you are working on.

    2. How do these advantages and disadvantages apply to the original VBA script?
Most of the work was already provided from this script. All we were asked to do was to make it more efficient. But having that code prewritten I did have to go through to make sure it would support the dataset we were given.
