using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using QuestPDF.Previewer;
using System.Data.Common;


namespace SalarySpecPdfCreator
{
    public partial class Form1 : Form
    {
        public static QuestPDF.Infrastructure.Image LogoImage { get; } = QuestPDF.Infrastructure.Image.FromFile("logo.png");

        private List<PaySlip> paySlipItems = [
                    new PaySlip {
                        Designation = "Timlön",
                        Quantity = 38,
                        Unit = "timmar",
                        APrice = 12.5,
                    },
                    new PaySlip {
                        Designation = "Semesterersättning",
                        Quantity = 38,
                        Unit = "timmar",
                        APrice = 100,
                    },
            new PaySlip {
                        Designation = "Semesterersättning",
                        Quantity = 38,
                        Unit = "timmar",
                        APrice = 100,
                    },
            new PaySlip {
                        Designation = "Semesterersättning",
                        Quantity = 38,
                        Unit = "timmar",
                        APrice = 100,
                    },
            new PaySlip {
                        Designation = "Semesterersättning",
                        Quantity = 38,
                        Unit = "timmar",
                        APrice = 100,
                    },
            new PaySlip {
                        Designation = "Semesterersättning",
                        Quantity = 38,
                        Unit = "timmar",
                        APrice = 100,
                    },
            new PaySlip {
                        Designation = "Semesterersättning",
                        Quantity = 38,
                        Unit = "timmar",
                        APrice = 100,
                    },
            new PaySlip {
                        Designation = "Semesterersättning",
                        Quantity = 38,
                        Unit = "timmar",
                        APrice = 100,
                    },
            new PaySlip {
                        Designation = "Semesterersättning",
                        Quantity = 38,
                        Unit = "timmar",
                        APrice = 100,
                    },
            new PaySlip {
                        Designation = "Semesterersättning",
                        Quantity = 38,
                        Unit = "timmar",
                        APrice = 100,
                    },
            new PaySlip {
                        Designation = "Semesterersättning",
                        Quantity = 38,
                        Unit = "timmar",
                        APrice = 100,
                    },
            new PaySlip {
                        Designation = "Semesterersättning",
                        Quantity = 38,
                        Unit = "timmar",
                        APrice = 100,
                    },
            new PaySlip {
                        Designation = "Semesterersättning",
                        Quantity = 38,
                        Unit = "timmar",
                        APrice = 100,
                    },
            new PaySlip {
                        Designation = "Semesterersättning",
                        Quantity = 38,
                        Unit = "timmar",
                        APrice = 100,
                    },
            new PaySlip {
                        Designation = "Semesterersättning",
                        Quantity = 38,
                        Unit = "timmar",
                        APrice = 100,
                    },
            new PaySlip {
                        Designation = "Semesterersättning",
                        Quantity = 38,
                        Unit = "timmar",
                        APrice = 100,
                    },
            new PaySlip {
                        Designation = "Semesterersättning",
                        Quantity = 38,
                        Unit = "timmar",
                        APrice = 100,
                    },
            new PaySlip {
                        Designation = "Semesterersättning",
                        Quantity = 38,
                        Unit = "timmar",
                        APrice = 100,
                    }

                    ];

        private Address fromAddress = new()
        {
            Name = "weteachitsweden AB",
            Street = "Rotegatan 8",
            ZipCode = "582 21",
            City = "Linköping",
            Email = "info@weteachit.se",
            WebAddress = "www.weteachit.se",
            Phone = "0142-345 84 52",
            MobileNumber = "0708-816 668",
            OrganisationNumber = "559209-9161"
        };

        private Address toAddress = new()
        {
            Name = "Kalle Kula",
            Street = "Storgatan 4",
            ZipCode = "587 52",
            City = "Linköping",
        };

        public Form1()
        {
            InitializeComponent();

            QuestPDF.Settings.License = LicenseType.Community;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Document.Create(container =>
            {
                container.Page(page =>
                {
                    page.Margin(30);

                    page.Header().Element(ComposeHeader);
                    page.Content().Element(ComposeContent);

                    page.Footer().Element(ComposeFooter);
                });
            }).GeneratePdf("hello world.pdf");
        }

        private void ComposeHeader(IContainer container)
        {
            int invoiceNumber = Random.Shared.Next(1_000, 10_000);

            container.Row(row =>
            {
                row.RelativeItem().Column(column =>
                {
                    column
                        .Item().Text($"Lönespec #{invoiceNumber}")
                        .FontSize(20).SemiBold().FontColor(Colors.Blue.Medium);

                    column.Item().Text(text =>
                    {
                        text.Span("Löneperiod: ").SemiBold();
                        text.Span($"{DateTime.Now.AddMonths(-1):d} till {DateTime.Now:d}");
                    });

                    column.Item().Text(text =>
                    {
                        text.Span("Utbetalningsdatum: ").SemiBold();
                        text.Span($"{DateTime.Now:d}");
                    });
                    column.Item().PaddingTop(10).Text(text =>
                    {
                        text.Span("Bankkontonummer: ").SemiBold();
                        text.Span($"3344-6 123456-7");
                    });
                });

                row.ConstantItem(175).Image(LogoImage);
            });
        }

        private void ComposeContent(IContainer container)
        {
            
            container.PaddingVertical(40).Column(column =>
            {
                column.Spacing(20);

                column.Item().Row(row =>
                {
                    row.RelativeItem().Component(new AddressComponent("Från", fromAddress));
                    row.ConstantItem(50);
                    row.RelativeItem().Component(new AddressComponent("Till", toAddress));
                });

                column.Item().Element(ComposeTable);

                column.Item().PaddingRight(5).AlignRight().Element(ComposeSubTable);

                
            });
        }

        private void ComposeTable(IContainer container)
        {
            var headerStyle = TextStyle.Default.SemiBold();

            container.Table(table =>
            {
                table.ColumnsDefinition(columns =>
                {
                    columns.ConstantColumn(25);
                    columns.RelativeColumn(3);
                    columns.RelativeColumn();
                    columns.RelativeColumn();
                    columns.RelativeColumn();
                    columns.RelativeColumn();
                });

                table.Header(header =>
                {
                    header.Cell().Text("#");
                    header.Cell().Text("Benämning").Style(headerStyle);
                    header.Cell().AlignRight().Text("Antal").Style(headerStyle);
                    header.Cell().AlignRight().Text("Enhet").Style(headerStyle);
                    header.Cell().AlignRight().Text("A-pris").Style(headerStyle);
                    header.Cell().AlignRight().Text("Belopp").Style(headerStyle);

                    header.Cell().ColumnSpan(6).PaddingTop(5).BorderBottom(1).BorderColor(Colors.Black);
                });

                

                foreach (var item in paySlipItems)
                {
                    var index = paySlipItems.IndexOf(item) + 1;

                    table.Cell().Element(CellStyle).Text($"{index}");
                    table.Cell().Element(CellStyle).Text(item.Designation);
                    table.Cell().Element(CellStyle).AlignRight().Text($"{item.Quantity}");
                    table.Cell().Element(CellStyle).AlignRight().Text($"{item.Unit}");
                    table.Cell().Element(CellStyle).AlignRight().Text($"{item.APrice:C}");
                    table.Cell().Element(CellStyle).AlignRight().Text($"{item.Amount:C}");

                }
                    static IContainer CellStyle(IContainer container) => container.BorderBottom(1).BorderColor(Colors.Grey.Lighten2).PaddingVertical(5);
            });
        }

        private void ComposeSubTable(IContainer container)
        {
            double totalAmount = paySlipItems.Sum(x => x.Amount);

            container.Table(table =>
            {
                table.ColumnsDefinition(columns =>
                {
                    columns.ConstantColumn(160);
                    columns.ConstantColumn(80);
                });

                table.Cell().Row(1).ColumnSpan(2).Text("Period").Bold();
                table.Cell().Row(2).Column(1).Text("Bruttolön");
                table.Cell().Row(2).Column(2).AlignRight().Text($"{totalAmount:C}");
                table.Cell().Row(3).Column(1).Text("Semesterersättning");
                table.Cell().Row(3).Column(2).AlignRight().Text($"{475:C}");

                double tax = totalAmount * 0.32;
                table.Cell().Row(4).Column(1).Text("Skatt");
                table.Cell().Row(4).Column(2).AlignRight().Text($"{tax:C}");

                table.Cell().Row(5).ColumnSpan(2).PaddingTop(5).BorderBottom(1).BorderColor(Colors.Black);
                table.Cell().Row(6).Column(1).Text("Utebetalt").Bold();
                table.Cell().Row(6).Column(2).AlignRight().Text($"{totalAmount - tax:C}").Bold();
            });
        }

        private void ComposeFooter(IContainer container)
        {
            container.Table(table =>
            {
                table.ColumnsDefinition(columns =>
                {
                    columns.ConstantColumn(120);
                    columns.RelativeColumn();
                    columns.RelativeColumn();
                    columns.RelativeColumn();
                    columns.RelativeColumn();
                    columns.RelativeColumn(1);
                });

                table.Header(header =>
                {
                    header.Cell().ColumnSpan(6).PaddingTop(5).BorderBottom(5).BorderColor(Colors.Black);

                    header.Cell().Text("Adress").SemiBold();
                    header.Cell().Text("");
                    header.Cell().AlignRight().Text("Org.nr").SemiBold();
                    header.Cell().AlignRight().Text("Tele").SemiBold();
                    header.Cell().AlignRight().Text("Web").SemiBold();
                    header.Cell().AlignRight().Text("E-post").SemiBold();

                });

                table.Cell().Row(1).Column(1).Text(fromAddress.Name).FontSize(8);
                table.Cell().Row(1).Column(3).AlignRight().Text(fromAddress.OrganisationNumber).FontSize(8);
                table.Cell().Row(1).Column(4).AlignRight().Text(fromAddress.MobileNumber).FontSize(8);
                table.Cell().Row(1).Column(5).AlignRight().Text(fromAddress.WebAddress).FontSize(8);
                table.Cell().Row(1).Column(6).AlignRight().Text(fromAddress.Email).FontSize(8);

                table.Cell().Row(2).Column(1).Text(fromAddress.Street).FontSize(8);
                table.Cell().Row(3).Column(1).Text($"{fromAddress.ZipCode} {fromAddress.City}").FontSize(8);
                table.Cell().Row(4).Column(1).Text("Godkänd för F-skatt");
                table.Cell().Row(4).ColumnSpan(5).AlignCenter().Text(text => {
                    text.CurrentPageNumber().FontSize(8);
                    text.Span(" / ").FontSize(8);
                    text.TotalPages().FontSize(8);
                });

                /*table.Footer(footer =>
                {
                    footer.Cell().Text("Godkänd för F-skatt");
                });*/

            });

        }

        public class AddressComponent : IComponent
        {
            private string Title { get; }
            private Address Address { get; }

            public AddressComponent(string title, Address address)
            {
                Title = title;
                Address = address;
            }

            public void Compose(IContainer container)
            {
                container.ShowEntire().Column(column =>
                {
                    column.Spacing(2);

                    column.Item().Text(Title).SemiBold();
                    column.Item().PaddingBottom(5).LineHorizontal(1);

                    column.Item().Text(Address.Name);
                    column.Item().Text(Address.Street);
                    column.Item().Text($"{Address.ZipCode} {Address.City}");

                    if (Address.Email != null) column.Item().Text($"E-post: {Address.Email}");
                    if (Address.Phone != null) column.Item().Text($"Tele: {Address.Phone}");
                    if (Address.MobileNumber != null) column.Item().Text($"Mobil: {Address.MobileNumber}");
                    if (Address.WebAddress != null) column.Item().Text($"Web: {Address.WebAddress}");
                    if (Address.OrganisationNumber != null) column.Item().Text($"Org.nr: {Address.OrganisationNumber}");
                });
            }
        }

        public class Address
        {
            public string Name { get; set; }
            public string Street { get; set; }
            public string ZipCode { get; set; }
            public string City { get; set; }
            public string? Email { get; set; } = null;
            public string? Phone { get; set; } = null;
            public string? MobileNumber { get; set; } = null;
            public string? OrganisationNumber { get; set; } = null;
            public string? WebAddress { get; set; } = null;
        }

        public class PaySlip
        {
            public string Designation { get; set; }
            public int Quantity { get; set; }
            public string Unit { get; set; }
            public double APrice { get; set; }
            public double Amount => Quantity * APrice; 
        }
    }
}
