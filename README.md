# ExcelDataComparator
Excel Data Comparator for comparison between different sheets with different row orders and different row quantities.

As Data Integrity becomes important part of daily work, identifying any unknown update, create and delete empowers data users to take precaution to avoid making mistakes.
Excel is used in most office work, however the Excel standard function simply compare exactly same cells in different sheets, which becomes inconvenient when the new sheet has different order, deleted or added rows. Hereby the Excel DataComparator project has been initiated.

This 1.1 version file provides basic function of data comparison.

It imports excel sheets from other files, and compare each cells of same rows in both old sheet, and new sheet then highlights the differences. The rows are identified by the Identify Column, hereby the order of the rows in each sheet can be different, and the 2nd sheet is allow to delete some old rows, and add new rows, which are also highlighted in both sheets.

In some circumstances multiple columns may needed to distinguish rows, only need to add the multiple column characters with comma “,”, and put the csv string as the Identify Column.
 
It would be great if you could provide some feedback, or any new function you prefer.
