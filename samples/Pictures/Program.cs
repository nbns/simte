using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Mime;
using simte.EPPlus;
using simte.Table;

namespace Pictures
{
    class Program
    {
        static void Main(string[] args)
        {
            using (var package = new Package())
            {
                var appRoot = AppContext.BaseDirectory.Substring(0,
                    AppContext.BaseDirectory.IndexOf("bin", StringComparison.Ordinal));


                var worksheet1 = package.AddWorksheet("newWorksheet-1");

                var image = Image.FromFile(Path.Combine(appRoot, "Images", "Lenna_(test_image).png"));
                worksheet1.Picture("nameOfPicture111", (1, 1), (2, 1), image);
                worksheet1.Picture("nameOfPicture222", (3, 3), image);


                var worksheet2 = package.AddWorksheet("newWorksheet-2");
                var endPos = worksheet2.Picture("Lenna", (1, 3), image);

                worksheet2.Text("Lenna or Lena is the name given to a standard test image widely used in the field of image processing since 1973.",
                    (endPos.Row + 3, 3),
                    opt => opt.Colspan(endPos.Col - 2).Rowspan(2).BackgroundColor(Color.Aqua)
                );

                var diagonalImage = Image.FromFile(Path.Combine(appRoot, "Images", "test-diagonal.png"));
                worksheet2.Table(new TableOptions {TopLeft = (endPos.Row + 6, 1)})
                    .AddRows(rowBuilder =>
                    {
                        rowBuilder
                            .Row.Column("1x2", opt => opt.Colspan(2)).ColumnRange(Enumerable.Range(1, 2).Select(n => $"1x1"))
                            .Row
                            .Column(diagonalImage, "imageInColumn-1", opt => opt.Colspan(2))
                            .Column(diagonalImage, "imageInColumn-2")
                            .Column(diagonalImage, "imageInColumn-3");
                    });


                package.Save("pictures.xlsx");
            }
        }
    }
}