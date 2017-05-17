# tsoyeow
This is a tool, written soley in .NET, that exports all resultsets from a SQL Server query to an Excel-compatible file (creates a ".xlsx" file that Excel 2007 and newer can open). 

It was purpose-written to quickly export large result sets efficiently (1 million rows and larger) to an Excel file. To that end it focuses on Excel-compatability, data type auto-detection, and speed; it offers few user-customizeable options (e.g., no support for creating Excel formulas or charts, etc.).
