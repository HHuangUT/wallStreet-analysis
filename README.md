# wallStreet-analysis
Challenge 2: refactoring wall street analysis code for optimized speed

# Overview of the project

>"Steve loves the workbook you prepared for him. At the click of a button, he can analyze an entire dataset. Now, to do a little more research for his parents, he wants to expand the dataset to include the entire stock market over the last few years. Although your code works well for a dozen stocks, it might not work as well for thousands of stocks. And if it does, it may take a long time to execute."

In this project, I refactor existing VBA code for optimized run times and scalability.

I used the provided VBA_Challenge.xlxm dataset, where the code's primary function is to scan wall street data for specific indexes over the course of a year that is entered by the user, and return 2 metrics:

1. Total Daily Volume

2. Return (%)

The values are then conditionally formated based on a 2 color scale: red for negative, green for positive.

# Results

In order to compare the two versions, timers built-in the macro were used to measure the runtime for each of the processes as shown below:

```
    Dim startTime As Single, endTime  As Single
    yearValue = InputBox("What year would you like to run the analysis on?")
    startTime = Timer  
```

*Note: the measured start time begins **after** the user interaction to prevent user delay/input from affecting the times*

When the macro completes it's process, the total runtime is shown as a message box (both code and 2018 UI examples shown below, respectively):


```
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
```
![VBA_Challenge_2018](https://user-images.githubusercontent.com/80546200/111920950-e9903380-8a5f-11eb-8a1f-dbd50c24299f.png)

The data is compiled and formated in the "All Stocks Analysis" sheet (Both 2017 and 2018 examples shown below):


![image](https://user-images.githubusercontent.com/80546200/111920968-12b0c400-8a60-11eb-8095-fda038f73c7c.png)
![image](https://user-images.githubusercontent.com/80546200/111921021-573c5f80-8a60-11eb-9eaf-ee75f2c61d7a.png)

The formatting helps point out that the selected indexes in 2018 showed *significantly worse* performance compared to 2017.

Because this is a data analytics class, 1 trial is not enough to claim that the refactored version is faster. In order to provide better evidence that the refactored version is faster, I ran 10 trials of each year on the comparable versions as shown below:

![image](https://user-images.githubusercontent.com/80546200/111921357-20ffdf80-8a62-11eb-9029-a08928591d88.png)
![image](https://user-images.githubusercontent.com/80546200/111921452-aa171680-8a62-11eb-90ec-2fbb299f514e.png)
```
  2017 Average Difference (sec):		-0.492
  2018 Average Difference (sec):		-0.484
```
# Summary

**Advantages and Disadvantages of Refactoring Code**

*Advantages:*
- Can build off of pre-existing framework for solution
- Save time and makes it easier to follow industry standard practices
- Can be modular; pick and choose which components of code to use

*Disadvantages:*
- Can be hard to integrate to a custom solution
- Pre-Existing framework may not be optimized for scalability
- Version control and lack of commenting will yield spaggheti code situations

**Original Code vs Refactored Code**

The two versions of code acomplish the same thing but go through different paths to do so.

>**Original Code:** scan each row, process data per index, output when complete with index, move to next index

>**Refactored Code:** scan each row, process data for all indexes in 1 spot, output when all rows scanned

**Benefit of Original Code:** The original is easy to follow (even without the comments) because it follows a path similar to how one would perform the analysis by hand.

**Benefit of Refactored Code:** By storing the data in 1 location for ALL the indexes, the reformatted code saves time over each iteration by not having to bounce back and forth to clean out/overwrite the old index data. This is directly shown by the improved runtimes:
```
  2017 Average Difference (sec):		-0.492
  2018 Average Difference (sec):		-0.484
```
In addition, by storing as an array that is initiallized at the beginning of the script, the user can easily edit the script to include more/less indexes, and/or refactor again for additional features- improving scalability.




