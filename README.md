# Stock-Analysis
Data Analytics Module 2 

## Overiew

  The purpose of this analysis was to familiarize myself with the VBA editor.  The challenge was to transform raw stock data into a useable Excel spreadheet that would display the daily volume and return of said stocks.  Once complete, the macros were refactored to clean up the code and improve its run time and clarity.  Then a side by side comparison was achieved, displaying the benefits of the refactored code.

## Results

  The original code and refactored code both returned the identical information and formatting.  The difference was in the time it took to process the code as shown below:

### *** Original Code Run Times for 2017 and 2018 data sets: ***

![Original Script 2017 Time](https://user-images.githubusercontent.com/111530580/187942022-64749e6c-c5bf-4cad-9377-35a1a16b6527.png)
![Original Script 2018 Time](https://user-images.githubusercontent.com/111530580/187942044-aaeae4e5-4602-45a6-b10a-2edcfa2e5756.png)

### *** Refactored Code Run Times for 2017 and 201 data sets.: ***

![VBA_Challenge_2017](https://user-images.githubusercontent.com/111530580/187942348-22fdd6af-ba80-45d2-b5dd-492504173b95.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/111530580/187942362-34a70544-36b1-4d04-bd82-61a314ef8d97.png)

  The images above show that for the original "dirty" code, the run times were 0.898 and 0.890 for 2017 and 2018 respectively.  The refactored code vastly improved the performace with a 2017 time of 0.164 second and a 2018 time of 0.160 seconds.

## Summary
  
  The advantages of refactoring the code were related to performance and cleanliness.  
  The performace advantages were obvious as the run times are roughly 5X faster.  This is a difference of only half a second or so with the relatively small data set used in the example.  however, should a larger data set be introduced, one would have more confidence in achieving timely results.
  The cleanliness helps by allowing future coders, including the author, to imrove on or modify it with greater understanding of its purpose.  Code with unnecessary steps would be harder and more time consuming to parse through, hindering future efforts to improve on or use its funcitons.
  The only drawback of refactoring is the time it takes to initially do it.  One may be reluctant to rewrite, or edit code that seemingly gets the job done.
  
  In this particular instance, refactoring the code shortened it from 262 to 142 lines.
