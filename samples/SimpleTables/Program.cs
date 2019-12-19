using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using simte;
using simte.EPPlus;
using simte.Table;

namespace SimpleTables
{
    public class Person
    {
        public string Name { get; }
        public int Age { get; }

        // ctor
        public Person(string name, int age)
        {
            Name = name;
            Age = age;
        }

        public static List<Person> GetPersons() =>
            new List<Person>()
            {
                new Person("Chris", 28),
                new Person("Dennis", 45),
                new Person("Sarah", 29),
                new Person("Karen", 47),
            };
    }

    class Program
    {
        static void Main(string[] args)
        {
            using (var package = new Package())
            {
                var worksheet = package.AddWorksheet("newWorksheet");

                // add some text
                worksheet
                    .Text("Simple Text", (1, 1), opt => opt.Colspan(2))
                    .Text("Simple Text", (3, 8), opt => opt.Colspan(2).FontBold(true).FontSize(12));

                // simple table 
                worksheet.Table(new TableOptions {TopLeft = (5, 2)})
                    .AddRows(rowBuilder =>
                        rowBuilder
                            .Row.Column( /*EmptyColumn*/).ColumnRange(Enumerable.Range(1, 3))
                            .Row.Column( /*EmptyColumn*/).Column("One").Column("Two").Column("Three")
                    );

                //see example: https://developer.mozilla.org/docs/Learn/HTML/Tables/Basics
                var pos = new Position(10, 2);
                worksheet.Table(new TableOptions {TopLeft = pos})
                    .AddRows(rowBuilder =>
                            rowBuilder
                                .Row.Column("Person", opt => opt.BackgroundColor(Color.Gray)).Column("Age", opt => opt.BackgroundColor(Color.Gray))
                                .Row.Column("Chris", opt => opt.BackgroundColor(Color.Gray)).Column(38)
                                .Row.Column("Dennis", opt => opt.BackgroundColor(Color.Gray)).Column(45)
                                .Row.Column("Sarah", opt => opt.BackgroundColor(Color.Gray)).Column(29)
                                .Row.Column("Karen", opt => opt.BackgroundColor(Color.Gray)).Column(47)
//                            .Row.Column("Karen2", mainOption)
                    );

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

                var tableBuilder = worksheet.Table(new TableOptions
                {
                    TopLeft = (28, 2),
                    Border = false
                });

                tableBuilder.AddRows(rowBuilder => rowBuilder.Row.Column().Column("Name").Column("Age").Column("FirstLetters", opt => opt.Width(20)));
                tableBuilder.AddRows(Person.GetPersons, rowBuilder =>
                {
                    rowBuilder.Row
                        .Skip(1)
                        .Select(x => x.Name, opt => opt.TextColorIf(x => x.Age > 37, Color.Aqua))
                        .Select(x => x.Age)
                        .Select(x => x.Name.Substring(0, 3).ToUpper())
                        .Apply();
                });

                // apply horizontal alignment
                worksheet.Table(new TableOptions() {TopLeft = (35, 2)})
                    .AddRows(rowBuilder =>
                        rowBuilder
                            .Row.Column(1, opt => opt.Colspan(2).HorizontalAlignment(HorizontalAlignment.Left))
                            .Row.Column(1, opt => opt.Colspan(2).HorizontalAlignment(HorizontalAlignment.Center))
                            .Row.Column(1, opt => opt.Colspan(2).HorizontalAlignment(HorizontalAlignment.Right))
                    );

                package.Save("simple-tables.xlsx");
            }
        }

        private static void mainOption(ColumnOptionsBuilder optionsBuilder) =>
            optionsBuilder.BackgroundColor(Color.Red);
    }
}