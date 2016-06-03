# Fast Excel Report Builder (FERB)

A tool to help quickly create simple OpenXML Excel reports.

## Overview
EPPlus is a powerful and thorough library for building Excel workbooks, which is why it is at the core of FERB. All that flexibility comes at the cost of complexity, though. I created FERB to make it quick and easy to generate simple reports in Excel. Initially, there will only be support for images, text blocks and tables.

Here's an example that builds up a report:

    WorkbookBuilder workbook = new WorkbookBuilder();
    IWorksheetBuilder worksheet = workbook.AddWorksheet("Account History");

    // Add logo
    var logo = worksheet.AddImage("Logo")
        .WithColumnSpan(6)
        .ScaledTo(75)
        .WithImage("some/path/to/logo.jpg");

    // Add report title
    var title = worksheet.AddTitleBar()
        .StartingAfter(logo)
        .StartingAt(2, 1)
        .WithColumnSpan(6)
        .WithStyleApplied(s =>
        {
            s.Bold = true;
            s.FontSize = 24;
        })
        .WithText("Account History");

    // Add Account Name
    var accountName = worksheet.AddTitleBar()
        .StartingAfter(title)
        .StartingAt(2, 1)
        .WithColumnSpan(6)
        .WithStyleApplied(s =>
        {
            s.Bold = true;
            s.FontSize = 16;
        })
        .WithText("Bob's Bait Shop");

    // Add event table
    var eventTable = worksheet.AddTable<Event>()
        .StartingAfter(accountName)
        .StartingAt(2, 1)
        .WithHeaderStyleApplied(s =>
        {
            s.Bold = true;
            s.SetBackgroundColor(Color.LightGray);
        })
        .WithColumn("Event Date", e => e.EventDate, x => x.Format = "g")
        .WithColumn("Event", e => e.EventType)
        .WithColumn("Initiator", e => e.UserName)
        .WithNestedTable(e => e.FieldChanges, t =>
            t.WithHeaderStyleApplied(s =>
            {
                s.Bold = true;
                s.SetBackgroundColor(Color.LightGray);
            })
            .WithColumn("Field Name", c => c.FieldName)
            .WithColumn("Original Value", c => c.OldValue)
            .WithColumn("Updated Value", c => c.NewValue)
        )
        .WithData(Events);

    // Add timestamp
    string createdOn = String.Format("Created on: {0:g}", DateTime.Now);
    worksheet.AddTitleBar()
        .StartingAfter(eventTable)
        .StartingAt(2, 1)
        .WithColumnSpan(6)
        .WithText(createdOn);
    
Just some highlights:
- The top-level class is the WorkbookBuilder. It provides:
	- The standard MIME type of the Excel document
	- The standard extension of the Excel document (.xslx)
	- Methods to create new WorksheetBuilders.
	- Save the Worksheet to a stream. You can even specify a template file and your report will be merged into it.
- The WorkSheetBuilder allows you to add items to the worksheets.
	- You can specify whether the columns should be auto-sized. This is the default.
	- By default, all items are added at the top of the report, overlapping each other.
	- You can use `StartingAfter` to indicate an item should appear after another.
	- You use `StartingAt` to indicate the number of rows and cells away from the starting point to start.
- With tables, you specify the columns and the data source separately.
	- Columns can be decorated with different styles, have formatting applied and be hidden based on a condition.
- With images, there is some support for scaling and stretching across cells.

## License
This is free and unencumbered software released into the public domain.

Anyone is free to copy, modify, publish, use, compile, sell, or
distribute this software, either in source code form or as a compiled
binary, for any purpose, commercial or non-commercial, and by any
means.

In jurisdictions that recognize copyright laws, the author or authors
of this software dedicate any and all copyright interest in the
software to the public domain. We make this dedication for the benefit
of the public at large and to the detriment of our heirs and
successors. We intend this dedication to be an overt act of
relinquishment in perpetuity of all present and future rights to this
software under copyright law.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
IN NO EVENT SHALL THE AUTHORS BE LIABLE FOR ANY CLAIM, DAMAGES OR
OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE,
ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR
OTHER DEALINGS IN THE SOFTWARE.

For more information, please refer to <http://unlicense.org>
