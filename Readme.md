# Coffee Orders Dashboard with Excel
## _Simple yet good enough_

Reference:
[Mo Chen Video](https://youtu.be/m13o5aqeCbM?si=UARzR0bPvWbIvG2J)

Doing excel is good and its pretty common for us, it also can used for work and creating awesome dashboard.

## Early Setup

Heres the early steps i did
1. Cleansing - Cleaning duplicates: some common files are often has duplicate rows and some empty / missing value. Simply select the identifier column and then choose `conditional formatting` -> `highlight duplicate values`, then go to `Data` tab and choose `remove duplicates`.
2. LOOKUPs: I'm using LOOKUPs for making a few better columns such as "Customer Name" and "Country". The data / value are in another sheet so i can get those from using "Custome ID" value from the current sheet.
3. Index match: Similar to LOOKUPs, i use this to get other value for new columns such as Roast type, unit price, and it size. The values are also in the different sheet so i select the headers for the reference in order to use the formula.
4. Multiple IFs: I need to do some multiple IFs formula to tell some better information. For example, there are "Coffee Type" column with the values are "Rob", "Lib", etc. So i use multiple IFs to tell that "Rob" means Robusta, "Lib" means Liberica, etc.
5. Date formatting: The "Order Date" columns format hasn't accurate because its just a classic number. So i change format for the column with dd-mmm-yyyy (or other) to become more realistic and i can use it later in pivot table.
6. Number formatting: Some numbers value like Prices need a currency type so its more valuable. To do that, i change the format to some currency.
7. Convert range to tables. I simply change the whole range to a table by press `ctrl + t`

| Formula | Function |
| ------ | ------ |
| LOOKUPs | Searches for a value in a range and returns a corresponding value from another range|
| Index Match | Finds position for a value and retrieves it at a position |
| Multiple IFs | Customize multiple specials condition and merge it |

Allright, those are the early steps. Next up, we're going to the pivot table section. Here we sort some information and generate the chart for visualization. Last, we make the dashboard in the different sheet.
## Pivot Tables
There are 3 information that i want to show in the dashboard
1. Total Sales over Time
2. Highest Sales Country
3. Top 5 Customer

So i make the pivot tables about those three
1. The first one is about Years vs Sales (USD). This shows the movement of the sales each year.
2. The second one shows the Sales vs Every Country. I sort it with the largest at the top so we can easily see and draw a conclusion
3. The last one is showing the most loyal customer by showing their sales vs the their name. I also sort it at Top 5 so we can see the top 5 customer by their sales.


## Dashboarding
Most steps are done and now we're ready to make a dashboard. Simply create a new sheet and create a rectangle for the title. Combine all graphs. To make it more dynamic, i add some Timeline and Slicer.
- I give Order Date timeline so we can see the different of the information for some times later.
- I add two slicers (roast name and coffee sizes) to see some unique POV from its size.
