# tsoyeow
This is a tool, written soley in .NET, that exports all resultsets from a SQL Server query to an Excel-compatible file (creates a ".xlsx" file that Excel 2007 and newer can open). 

It was purpose-written to quickly export large result sets efficiently (1 million rows and larger) to an Excel file. To that end it focuses on Excel-compatability, data type auto-detection, and speed; it offers few user-customizable options (e.g., no support for creating Excel formulas or charts, etc.).

Features:
* Quickly export large volumes of data from a SQL Server query to an Excel file (or files, optionally)
* Data type autodetection without data loss (e.g., numbers are written as Excel number types when safe to do so, dates are auto-detected, etc.)
* Multiple result sets can be written to new tabs or new workbooks
* Supports multiple SQL Server authentication types
* Multi-threaded design to let you quickly see what's happening in the UI
* Entirely .NET code
