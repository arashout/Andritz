# ANDRITZ VBA Macros
## Get Material Information
This program allows users to quickly pull information about a material given an SAP number
Probably my most popular program at the office because everyone gets some use out of it.
For example:
Project Managers or Product Managers who need to create quotes for customers previously had to manually look up each SAP number to get the most recent price which is super repetitive and time consuming because there would be upwards of 30 SAP numbers to look up.
Now they simply have to select the cells with SAP numbers and run my program which saves them hours.
It's so popular that it has been distributed throughout Andritz not just at the Delta Service Center

## Get Routing
Originally designed by my friend the week before he left. I redesigned it in a much more modular fashion using OO principles, which made it much more extensible (If in the future I needed to add features) and easier to read.

## Set Routing
Often times PMs keep track of there routings for each job on an Excel file but later on when the job is about to start they have to re-enter this information into SAP which means they are duplicating work perhaps even tripling it because entering information into SAP can be slow (which is why they keep it in Excel sheets until they have to update SAP). 
I've created this macro to allow them to quickly push changes in there operations to SAP with a little extra effort on their part. 
I have implemented it so that users can change routings and production orders

### Background Information
* OPERATION = The atomic level job that shop employees will perform on a part e.g. Face off part after welding
* SEQUENCE = A list of operations in the ORDER they should be done
* ROUTINGS = list of sequences which can be parallel (being done at the same time) or standand (Have to be done right after each other)

## Set BOM
Very similar to "Set Routing" but probably a little simpler. When project managers want to update the items in the BOM of a project it was a repititive process. I've stream-lined it with this macro. Also note that this works with template BOMs as well as production order BOMs

## Help Functions
This is a big module that contains all the utility subs and functions I've collected. I often use small bits of them in my macros 

## SAP Functions
These functions allow interface with SAP Gui, allowing me to act like a user that clicks on stuff
