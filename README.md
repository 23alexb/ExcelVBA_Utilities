# ExcelVBA_Utilities
Excel Macros that may be useful for office and data analysis/cleaning applications.

Macros:
 - ColorCodeRows:
   - Macro to color-code the used range (or selected range) of a worksheet
   - Colors alternate whenever the value in a row specified by the user changes
   - Useful for easy viewing of ordered categorized data
 - ConvertCellValueType:
   - Four macros to convert cells and cell values to a specified format
   - Types are text, date, integer, or decimal
   - Used in instances where formatting options in the ribbon fail to convert the 
     underlying cell value when the format is changed
 - DataIn_DataOut:
   - Macro to add VLOOKUP formulas to "merge" data between a data table and a source 
     range with a shared key
   - Macros to copy a column of keys, adding a prefix and suffix, that can be pasted
     into the WHERE (or other) clause of a SQL query
   - Used to copy a set of keys into a query, then merge the query output back into
     the source worksheet
