# Vaccine-Router

This was designed to assist in assigning vaccine packages to routes based on their zip code.

The [Router.xlsm](./Router.xlsm) can be edited with route numbers on top with the desired zip codes for that route in the column below. 

When the Route button is clicked it will open up a file explorer for you to open the CSV containing the vaccine info. Note: this edits the CSV you open, if you would like an unalterd version be sure to make a backup first.

Once a file is selected it will automatically remove columns that are not needed and sort the colums to give a consistent format reguardless of the columns starting order.

It will then look at the first cell in a column and use that as a route number and use the zip codes below to assign route numbers. 

This could be made robust, and add automatic stop counts and other information. However, due to lack of interest from the company I have not continued development and continue to use this in its current state as it achieves its main purpose in saving time by not needing to manually look at each zip code and enter route numbers.
