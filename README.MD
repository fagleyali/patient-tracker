# How to read and write from Excel file

#### This file helps to understand and shows how to read and write excel files in the node.js. 

#### I used a Patient List consisting of 20 sheets, but this application reads only 8 sheets depending on sheets name. each of 8 sheets was iterated based on First Visit Date Column value (which dates match with the requested month by client) and write to another sheet for generating Marketer Report for the selected month.  