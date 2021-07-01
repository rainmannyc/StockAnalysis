#**VBA Stock Analysis**

##**Overview**

The "Stock Analysis VBA" macros were originally created for our client Steve, who requested a way to collect and process specific stock data in which we acquired from excel data sheets. Our client Steve required a quick and simplified way to calculate and process specific data from the stock tickers listed in the data sheets. 

##**The Macro**

The Macro we created for Steve analyzed all the data on the sheets provided and outputs the results he requested. The "Total Daily Volume" and "Return" were calculated by year and organized by stock tickers and outputs onto a separate data sheet for Steve called "All Stocks Analysis".

The analyzation process was simplified for Steve by creating a button he can simply click and then input the year he would like to view the data for in an input box which pops up upon clicking the button.

By using our macro, he was able to clearly see which stocks resulted with positive or negative outcomes throughout the years he inputs. For this instance, Steve can clearly see stock data such as stock ticker "RUN" having an increase of 78.5% on the returns from year 2017 to 2018. And although stock ticker "ENPH" dropped 47.6%, the return was still in the positives due to the significant increase in volume. In which Steve can now clearly see indicated by the color formats. Steve may now proceed to making his assessment and decisions on how he would like to move forward more clearly, pulling up the necessary data he requires at the click of a button!

###**THE Codes**

For our codes, we used arrays, conditional statements, dimensions and conditional formatting to create the macro.

The original code contained a "Nested For Loop" which caused the process of iterating over the rows using the conditional statements to take a longer period of time. As shown in the images linked below:

'IMG: of nested loop in DQ vs lack of in refactored

This is the effect of the code having to loop over all the rows while also being the output source for the current data in the process.

####**Original VBA Codes vs Refactored VBA Codes**

The advantage compared to the original code is in the processing time required to run the code. It is significantly faster as shown in the recorded times: 

vba_challenge_2017.png & vba_challenge_2018.png 
                    VS 
Refactoredvba_challenge_2017.png & Refactoredvba_challenge_2018.png 

This is due to the code allowing the process initiated to process the same results with a different approach. This would be advantageous if our client would like to run the stocks for multiple years, with multitude of stock data to process as opposed to the slower run time in the original script. I believe this may also be beneficial if they have different devices with lower processing power across different platforms which are also able to run Excel and VBA (i.e Mobile Phone, Tablet, or older computer models). 

However, on the downside, if there are certain edits to the code, for example if our client would like to make certain changes or updates, we would have to review and revise the code again. As well as testing the code for the same or higher efficiency each time. This also means each time there would also be a possibility of bugs arising to be solved as well.

####**The Advantages of Refactoring Codes**

One of the advantages of refactoring is the possibility of using less processing power as well as speeding up the process. In essence, doing more with less. Refactoring the code also allows a clearer view for different perspectives from other programmers. The process may enable them to approach or solve certain issues with different methods. There may be multiple ways to solve one particular issue.

As mentioned earlier, although bugs may arise from refactoring, the process of refactoring may also help other peers troubleshoot issues in the codes that are repetitive or unnecessary to achieve the same result. Refactoring may "clean up" the code visually and operationally as you can also see in images linked above.

####**The Disadvantages of Refactoring Codes**

On the other hand, refactoring in general can be very time consuming. Thus, if the client is required to make constant edits or updates, there can be many issues that arise. Especially if there is a deadline. Changing the method or approach to solving an issue for one problem, may not be the same for the different issues you may encounter. There may be a deadline and depending on your resources and time frame, this may prove costly with a negative outcome. I believe it is also possible it could take the same amount of time or longer than writing a new code for the issue as well.

#####**Please note: The VBA Challenge code is located in "Module 3" in the VB file of VBA_Challenge.xslm. I have also attached an additional VBA Macro file titled "Mod2SubRFactorExplained.vb" in the repository with further explanations of specific lines of code. There is also an additional button created on the "All Stocks Analysis" worksheet for easier access to view the difference of speed in the original versus the refactored macro.**

