# Stock Analysis

## Analysis of Stocks for Steve

## Purpose
Provide analysis of Steve's stock data from 2017 and 2018. Steve would like to advise his parents which stock options are most likely to be profitable. With VBA, we can select a specific year's data in Excel easily to quickly provide comparisons. Refactoring the VBA meant a faster result when the macro was run.

## Results

After completion of the analysis, it is determined that overall stocks performed better in 2017 than in 2018. Results are shown below.
**The Results of the Analysis for 2017:**

![image](https://user-images.githubusercontent.com/95710184/148834877-87007e6b-e818-49f6-b9f4-a630921fb29e.png)

**The Results of the Analysis for 2017:**

![image](https://user-images.githubusercontent.com/95710184/148834890-1a5abad6-fb4c-4c76-a095-5b637acae790.png)


**Using VBA code for a timer, the time it takes to run was able to be captured.
     
     'Start timer
      Dim startTime As Single
      Dim endTime  As Single
      startTime = Timer
      
      'End Timer
      endTime = Timer

**Intially, the VBA code ran in:**

0.7265625 seconds for 2017

![image](https://user-images.githubusercontent.com/95710184/148707721-ff422f6a-7268-4aa8-ab2a-6e55f2979e2e.png)

0.734375 seconds for 2018

![image](https://user-images.githubusercontent.com/95710184/148707745-e632e317-e043-4511-98af-827de9d8e93b.png)

**Following refactoring, the VBA code ran in:**

0.25 seconds for 2017

![image](https://user-images.githubusercontent.com/95710184/148707788-80ec5760-4a95-4269-9990-20e9a166ea90.png)

0.125 seconds for 2018

![image](https://user-images.githubusercontent.com/95710184/148707781-7d678207-ae2a-4175-b387-f4a3c09c77d1.png)


**Refactoring improved the time taken to run the code.**

##Summary

**1) What are the advantages or disadvantages of refactoring code?**
The advantages of refactoring the code include increased efficiency in time to run the code and space required to save the code. Additionally, there are opportunities to improve the appearance and readability with refactoring by cleaning up anything unnecessarily complex. Refactored code is also more scalabale and potentially easier to reuse in the future. However, the time saving advantages may strongly favor very large data sets as the time saving is minimal on a data set of this size. If the code is working well, refactoring may not make sense, especially if deadlines are involved.

**2) How do these pros and cons apply to refactoring the original VBA script?**
Refactoring did allow me to clean up some areas of my VBA code that were not written as well as they could have been. However, the code was working and I actually introduced an error in refactoring that took me several hours to uncover. In switching from using totalVolume to totalVolumes (with an "s"), I overlooked one place in the refactored code it was an easy mistake to omit the "s", which resulted in a bug I had to fix. The cleaned up code is much easier to read, so I could see that it is helpful in some scenarios to iterate through the refactoring process.
