# tsoyeow
This is a tool, written in .NET, that exports results from SQL Server queries to an Excel-compatible file (creates a ".xlsx" file that modern Excel can open). 

It was written to quickly export large result sets efficiently (1 million rows and larger) to an Excel file. It's focused on Excel-compatability, data type auto-detection, and speed; it offers very few user-customizable options (e.g., no support for creating Excel formulas or charts, etc.).

Features:
* Quickly export large volumes of data from a SQL Server query to an Excel file (or files, optionally)
* Data type autodetection without data loss (e.g., numbers are written as Excel number types when safe to do so, dates are auto-detected, etc.)
* Multiple result sets can be written to new tabs or new workbooks
* Supports multiple SQL Server authentication types
* Multi-threaded design
* All .NET code
