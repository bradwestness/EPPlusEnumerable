# EPPlusEnumerable
Easily create multi-worksheet Excel documents from any .NET object collection.

## Usage

Let's say you're working on an ASP.NET web app and want to create a report of all users and orders.

```csharp
public ActionResult DownloadReport()
{
    var data = new List<IEnumerable<object>>();
    
    using (var db = new SampleDataContext())
    {
        data.Add(db.Users.OrderBy(x => x.Name).ToList());

        foreach(var grouping in db.Orders.OrderBy(x => x.Date).GroupBy(x => x.Date.Month))
        {
            data.Add(grouping.ToList());
        }
    }
    
    var bytes = Spreadsheet.Create(data);
    return File(bytes, "application/vnd.ms-excel", "MySpreadsheet.xlsx");
}
```

That will give you a nicely-formatted Excel spreadsheet with tabs for both "Users" and "Orders," like so:

![output](https://raw.githubusercontent.com/bradwestness/EPPlusEnumerable/master/output.png)

There's also a `SpreadsheetLinkAttribute` class which you can use to generate links between tabs on your spreadsheet.

```csharp
[DisplayName("Orders"), SpreadsheetTableStyle(TableStyles.Medium16)]
public class Order
{
    public int Number { get; set; }

    [SpreadsheetCellStyle(bold: true, backgroundHtmlColor: "#00cc00", fontHtmlColor: "lime")]
    public string Item { get; set; }

    [SpreadsheetLink("Customers", "Name")]
    public string Customer { get; set; }

    [DisplayFormat(DataFormatString = "{0:c}")]
    public decimal Price { get; set; }

    [SpreadsheetTabName(FormatString = "{0:MMMM yyyy}")]
    public DateTime Date { get; set; }
}
```

In this example, the "Customer" values in the Orders tab will be linked to the corresponding Customers tab row where the Name is equal to the value of the Order object's Customer property.

In addition, the `SpreadsheetTabName` attribute specifies that the value of the given property should be used to name the tab in Excel. The first value of the collection will be used to generate the name of the tab. You can optionally specify a `FormatString` as above.

![links](https://raw.githubusercontent.com/bradwestness/EPPlusEnumerable/master/links.png)

There's also a `SpreadsheetExclude` attribute you can use if you don't want to include a column for a given property in the generated spreadsheet.

P.S. `DisplayName` or `Display(Name="")` attributes will be used for worksheet names if used on the class, or column headers if used on a property.
