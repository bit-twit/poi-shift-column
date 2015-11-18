POI SHIFT COLUMN
=====================

Simulates Excel's facility to "insert column left".
Tries to keep the formulas in check:
- if the column moved right had formulas in it that referred to the given column, it updates those formulas

## RUN ##
* just do mvn clean test
	* modified excel file is saved in out.xlsx
