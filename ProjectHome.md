Tsoyeow is a pure c# solution for writing Excel-compatible files from a Sql Server query. Unlike other approaches, it uses a streaming approach to writing Excel files, enabling creation of Excel files from large query datasets without exhausting memory.

There are 3 projects contained within the solution:

  * The ExcelXmlWriter back-end, which processes a query into Excel 2007/2010 Xlsx-compatible files.
  * The ExcelXmlQueryResults front-end, which handles the GUI and provides an example implementation of how to call ExcelXmlWriter.
  * The ExcelXmlWriterNTest project, which provides a suite of unit tests to enforce functionality.

Tsoyeow is released under the Microsoft Public License.