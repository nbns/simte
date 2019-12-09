# simte

# Features

simte is a library that you can add in to your project that will generate excel spreadsheets with tables in a similar way like in HTML.

Examples usage:

```csharp
using (var package = new Package())
{
    var worksheet = package.AddWorksheet("newWorksheet");

    // add some text
    worksheet
        .Text("Simple Text", (1, 1))
        .Text("Simple Text", (3, 8), opt => opt.Colspan(2).FontBold(true).FontSize(12));

    worksheet.Table(new TableOptions() {TopLeft = (18, 2)})
        .AddRows(rowBuilder =>
            rowBuilder
                .Row.Column(opt => opt.Colspan(3)).Column("Subject").Column("Object")
                .Row.Column("Singular", opt => opt.Rowspan(5)).Column("Person", opt => opt.Colspan(2)).Column("I").Column("me")
                .Row.Column("2nd Person", opt => opt.Colspan(2)).Column("you").Column("you")
                .Row.Column("3rd Person", opt => opt.Rowspan(3)).Column("m").Column("me").Column("him")
                .Row.Column("w").Column("she").Column("her")
                .Row.Column("o").Column("it").Column("it")
                .Row.Column("Plural", opt => opt.Rowspan(3)).Column("Person", opt => opt.Colspan(2)).Column("we").Column("us")
                .Row.Column("2nd Person", opt => opt.Colspan(2)).Column("you").Column("you")
                .Row.Column("3nd Person", opt => opt.Colspan(2)).Column("they").Column("them")
        );
}
```

For more information see the samples
