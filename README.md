# VBA of Wall Street
Module 2 Challenge
## Overview of Project
This project was intended to provide a springboard for the heavier coding languages by introducing concepts such as For loops and their operation and interaction with one another, iterators and when/how to use them, and the formatting of data by accessing it's properties via command while in an active subroutine.
### Purpose
The purpose of this project is to drill into the basic framework of coding, with specific emphasis the logic of loops, defining variables using "Dim", and refactoring existing code to run more efficiently.
## Analysis and Results
The initial module was a crash course in code basics. VBA is a solid all-around trainer, as it introduces universal concepts to a learning coder in a manner that is clean, easy to read, and presented in a format that allows for quick editing and troubleshooting when something inevitably goes wrong. The power of VBA lies in it's reuseability and adaptation to larger data sets, as seen in the .xlsm file accompanying this README. If the Client were to upload another data sheet for 2019, 20, and so on, as long as the ticker data was the same as what we are already working with (and as it tends to stay for long term investment portfolios), we could infer that the data would populate just the same way as it does now in the "All Stocks Analysis" worksheet:

<img width="279" alt="Screenshot 2022-11-13 at 6 24 46 PM" src="https://user-images.githubusercontent.com/116296092/201556631-495b55b3-d75d-4fa3-9d3e-c0e9a1270eb6.png">

There are no doubt advantages to refactoring code to run more efficiently and be read easier, but on a set as small as this with the backstory it has I don't know how effective it is. The initial block of code from green_stocks (the original file) is less cluttered, but also relies more on nested loops and variables talking to each other. When dismantled, things stop working, and can be harder to track:

<img width="469" alt="Screenshot 2022-11-13 at 6 30 12 PM" src="https://user-images.githubusercontent.com/116296092/201557107-84c9d109-a742-4f7f-bb3f-511106012d01.png">

While this is easier to look at than the dense block that comprises VBA_Challenge (which will be shown shortly), it's less adaptable to new stocks. For example, the output block for green_stocks uses hard values, and thus is less flexible than the refactor in VBA_Challenge:

<img width="343" alt="Screenshot 2022-11-13 at 6 32 57 PM" src="https://user-images.githubusercontent.com/116296092/201557291-9648b879-c5d3-4841-af18-bbbcd90c5592.png">

vs:

<img width="451" alt="Screenshot 2022-11-13 at 6 33 25 PM" src="https://user-images.githubusercontent.com/116296092/201557319-525b835b-e7c1-43ae-90f5-52771f85fa83.png">

While the difference is subtle, the key ingredient is that the refactored code is able to loop through the output worksheet after all the values have been collected, avoiding a worksheet activation that wastes time, while the original is forced to use hard references and "deposit" on every loop. Multiple For loops that contain related processes before moving to the next step make the computer's job easier, and this is further shown in the full script:

<img width="620" alt="Screenshot 2022-11-13 at 6 40 02 PM" src="https://user-images.githubusercontent.com/116296092/201557822-3c4bb54e-79a0-4bd2-ab79-07c2f74497f0.png">

The arrays up top provide a place to store the values for all the stocks that we want to analyze while the Challenge Sub is running, as opposed to the green_stocks Sub needing to place individual values on the output worksheet every time prior to looping through again. The multiple For loops break the code up into more contained operations, and allows the efficiency of the code to increase by orders of magnitude.

Yes, the refactored code is a bit more complicated in it's construction, but far more efficient. This is proven by the speedtests run immediately prior to beginning the refactor, as shown below:

<img width="261" alt="VBA_Challenge_2018_MODULE" src="https://user-images.githubusercontent.com/116296092/201558185-936b1637-6796-4e69-85f7-64e3135a8fcb.png">

and then run again immediately after:

<img width="261" alt="VBA_Challenge_2018_REFACTOR" src="https://user-images.githubusercontent.com/116296092/201558205-66555972-386d-4b7d-b8b1-d6071b3ea5b8.png">

As explained by the module and the instructors, the Macro takes longer to run on the first execute than it does on subsequent executes, and this is shown by the REFACTOR_II screenshot:

<img width="263" alt="VBA_Challenge_2018_REFACTOR_II" src="https://user-images.githubusercontent.com/116296092/201558339-df1bfcf1-cb90-4c7d-842c-67739e1bf54b.png">

In general, even on this small of a scale, refactoring to move more efficiently is always best practice. While the difference is practically irrelevant to us in the case of this Challenge, if we were to consider an industrial production environment, and datasets that are scaled up however many times and constantly changing, that % difference in speed - and certainly computing efficiency - are going to be felt much more keenly than they are in this practice. It's harder to do and requires a bit more tinkering, but once it's set up it could be a perfect fire and forget utility (which is why Microsoft disables it by default, is it not?).

There is a disadvantage to refactoring this VBA in that it is more complicated than simply writing a "straightforward" Macro as done in the case of green_stocks. I know that personally it was easier to write green_stocks' Macro than it was to write the refactor, but the results don't lie - the refactored code is more efficient, and ultimately that is the goal. It tested my abilities, but that's why it was assigned. With how small the dataset is, it would likely work just fine in it's original form for it's intended use of small-group stock analysis for a private investor (parents of the Client), but if we can make it more efficient and easier to run then why wouldn't we?
