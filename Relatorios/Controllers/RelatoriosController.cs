using Microsoft.AspNetCore.Mvc;
using PuppeteerSharp;
using PuppeteerSharp.Media;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;

namespace Relatorios.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class RelatoriosController : ControllerBase
    {
        [HttpGet("grupo-economico")]
        public IActionResult GrupoEconomico()
        {
            QuestPDF.Settings.License = LicenseType.Community;
            var titular = "COLGATE-PALMOLIVE COMPANY";
            var escritorio = "LUIZ LEONARDOS & ADVOGADOS";
            var dataHoje = DateTime.Now.ToString("dd/MM/yyyy");

            var empresas = new List<(string Codigo, string Nome, string Data)>
            {
                ("16118","ARCOM S/A.","15/09/2010"),
                ("22901","AROMATERAPIA INDÚSTRIA E COMÉRCIO LTDA.","15/09/2010"),
                ("09590","C-P INDÚSTRIA E COMÉRCIO LTDA.","15/09/2010"),
                ("01253","COLGATE PALMOLIVE LTDA.","15/09/2010"),
                ("20074","COLGATE-PALMOLIVE CHILE S.A.","15/09/2010"),
                ("00843","COLGATE-PALMOLIVE COMPANY","15/09/2010"),
                ("36104","COLGATE-PALMOLIVE EUROPE SARL","29/12/2011"),
                ("31496","COLGATE-PALMOLIVE INVESTMENTS (BVI)","15/09/2010"),
                ("06361","COLGATE-PALMOLIVE S.p.A","15/09/2010"),
                ("08939","COMÉRCIO DE ESCOVA ORAL LTDA.","15/09/2010"),
                ("12362","DALTEX INDUSTRIAL LTDA","15/09/2010"),
                ("15394","FISBRA SERVIÇOS EM CONSULTORIA LTDA.","15/09/2010"),
                ("14292","HILL'S PET NUTRITION, INC.","15/09/2010"),
                ("11271","INDÚSTRIA E COMÉRCIO SANTA THEREZA LTDA.","15/09/2010"),
                ("05911","INDÚSTRIAS GESSY LEVER LTDA.","15/09/2010"),
                ("11664","KOLYNOS DO BRASIL LTDA","15/09/2010"),
                ("11124","ODOL SOCIEDAD ANONIMA INDUSTRIAL Y COMERCIAL","15/09/2010"),
                ("15471","RALSTON PURINA COMPANY","15/09/2010"),
                ("31602","TEC IMPORTS IMPORTAÇÃO E EXPORTAÇÃO LTDA.","15/09/2010"),
                ("07236","THE MENNEN COMPANY","15/09/2010"),
                ("12108","UNILEVER N.V.","15/09/2010"),
                ("08520","WYETH","15/09/2010"),
                ("00160","WYETH HOLDINGS CORPORATION","15/09/2010"),
                ("16118","ARCOM S/A.","15/09/2010"),
                ("22901","AROMATERAPIA INDÚSTRIA E COMÉRCIO LTDA.","15/09/2010"),
                ("09590","C-P INDÚSTRIA E COMÉRCIO LTDA.","15/09/2010"),
                ("01253","COLGATE PALMOLIVE LTDA.","15/09/2010"),
                ("20074","COLGATE-PALMOLIVE CHILE S.A.","15/09/2010"),
                ("00843","COLGATE-PALMOLIVE COMPANY","15/09/2010"),
                ("36104","COLGATE-PALMOLIVE EUROPE SARL","29/12/2011"),
                ("31496","COLGATE-PALMOLIVE INVESTMENTS (BVI)","15/09/2010"),
                ("06361","COLGATE-PALMOLIVE S.p.A","15/09/2010"),
                ("08939","COMÉRCIO DE ESCOVA ORAL LTDA.","15/09/2010"),
                ("12362","DALTEX INDUSTRIAL LTDA","15/09/2010"),
                ("15394","FISBRA SERVIÇOS EM CONSULTORIA LTDA.","15/09/2010"),
                ("14292","HILL'S PET NUTRITION, INC.","15/09/2010"),
                ("11271","INDÚSTRIA E COMÉRCIO SANTA THEREZA LTDA.","15/09/2010"),
                ("05911","INDÚSTRIAS GESSY LEVER LTDA.","15/09/2010"),
                ("11664","KOLYNOS DO BRASIL LTDA","15/09/2010"),
                ("11124","ODOL SOCIEDAD ANONIMA INDUSTRIAL Y COMERCIAL","15/09/2010"),
                ("15471","RALSTON PURINA COMPANY","15/09/2010"),
                ("31602","TEC IMPORTS IMPORTAÇÃO E EXPORTAÇÃO LTDA.","15/09/2010"),
                ("07236","THE MENNEN COMPANY","15/09/2010"),
                ("12108","UNILEVER N.V.","15/09/2010"),
                ("08520","WYETH","15/09/2010"),
                ("00160","WYETH HOLDINGS CORPORATION","15/09/2010"),
                ("16118","ARCOM S/A.","15/09/2010"),
                ("22901","AROMATERAPIA INDÚSTRIA E COMÉRCIO LTDA.","15/09/2010"),
                ("09590","C-P INDÚSTRIA E COMÉRCIO LTDA.","15/09/2010"),
                ("01253","COLGATE PALMOLIVE LTDA.","15/09/2010"),
                ("20074","COLGATE-PALMOLIVE CHILE S.A.","15/09/2010"),
                ("00843","COLGATE-PALMOLIVE COMPANY","15/09/2010"),
                ("36104","COLGATE-PALMOLIVE EUROPE SARL","29/12/2011"),
                ("31496","COLGATE-PALMOLIVE INVESTMENTS (BVI)","15/09/2010"),
                ("06361","COLGATE-PALMOLIVE S.p.A","15/09/2010"),
                ("08939","COMÉRCIO DE ESCOVA ORAL LTDA.","15/09/2010"),
                ("12362","DALTEX INDUSTRIAL LTDA","15/09/2010"),
                ("15394","FISBRA SERVIÇOS EM CONSULTORIA LTDA.","15/09/2010"),
                ("14292","HILL'S PET NUTRITION, INC.","15/09/2010"),
                ("11271","INDÚSTRIA E COMÉRCIO SANTA THEREZA LTDA.","15/09/2010"),
                ("05911","INDÚSTRIAS GESSY LEVER LTDA.","15/09/2010"),
                ("11664","KOLYNOS DO BRASIL LTDA","15/09/2010"),
                ("11124","ODOL SOCIEDAD ANONIMA INDUSTRIAL Y COMERCIAL","15/09/2010"),
                ("15471","RALSTON PURINA COMPANY","15/09/2010"),
                ("31602","TEC IMPORTS IMPORTAÇÃO E EXPORTAÇÃO LTDA.","15/09/2010"),
                ("07236","THE MENNEN COMPANY","15/09/2010"),
                ("12108","UNILEVER N.V.","15/09/2010"),
                ("08520","WYETH","15/09/2010"),
                ("00160","WYETH HOLDINGS CORPORATION","15/09/2010"),
            };

            var pdf = Document.Create(doc =>
            {
                doc.Page(page =>
                {
                    page.Size(PageSizes.A4);
                    page.Margin(56);
                    page.DefaultTextStyle(s => s.FontSize(11));

                    page.Header().Column(h =>
                    {
                        h.Item().Row(r =>
                        {
                            r.RelativeItem().Text(escritorio).FontSize(11).SemiBold();
                            r.ConstantItem(100).AlignRight().Text(dataHoje).FontSize(11);
                        });

                        h.Item().AlignRight().Text(t =>
                        {
                            t.Span("Página ").FontSize(11);
                            t.CurrentPageNumber().FontSize(11);
                            t.Span(" de ").FontSize(11);
                            t.TotalPages().FontSize(11);
                        });

                        h.Item().PaddingTop(2).Column(tituloCol =>
                        {
                            tituloCol.Item()
                               .Border(2)
                               .BorderColor(Colors.Black)
                               .Background(Colors.White)
                               .Padding(1)
                               .AlignCenter()
                               .Text($"Grupo Econômico {titular}")
                               .FontSize(12)
                               .Bold();

                            tituloCol.Item().PaddingTop(12);
                        });
                    });


                    page.Content().Column(col =>
                    {

                        col.Item().PaddingTop(12);
                        col.Item().Table(table =>
                        {
                            table.ColumnsDefinition(c =>
                            {
                                c.RelativeColumn(3);   
                                c.ConstantColumn(90); 
                            });

                            table.Header(h =>
                            {
                                h.Cell().Border(2).Padding(1).Text("Empresas Associadas").Bold().AlignCenter();
                                h.Cell().Border(2).Padding(1).Text("Data").Bold().AlignCenter();
                            });

                            foreach (var e in empresas)
                            {
                                table.Cell().Border(2).Padding(8).Text($"{e.Codigo} - {e.Nome}");
                                table.Cell().Border(2).Padding(8).AlignRight().Text(e.Data);
                            }
                        });


                        col.Item().PaddingTop(12);

                        col.Item()
                               .Border(1)
                               .BorderColor(Colors.Black)
                               .Background(Colors.White)
                               .Padding(8)
                               .AlignCenter()
                           .Text($"Total: {empresas.Count}")
                           .FontSize(10).Bold();


                    });
                });
            }).GeneratePdf();

            return File(pdf, "application/pdf", "relatorio-grupo-economico.pdf");
        }

        [HttpGet("grupo-economico-puppeteer")]
        public async Task<IActionResult> GrupoEconomicoPuppeteer()
        {
            var titular = "COLGATE-PALMOLIVE COMPANY";
            var escritorio = "LUIZ LEONARDOS & ADVOGADOS";
            var dataHoje = DateTime.Now.ToString("dd/MM/yyyy");

            var empresas = new List<(string Codigo, string Nome, string Data)>
    {

                ("16118","ARCOM S/A.","15/09/2010"),
                ("22901","AROMATERAPIA INDÚSTRIA E COMÉRCIO LTDA.","15/09/2010"),
                ("09590","C-P INDÚSTRIA E COMÉRCIO LTDA.","15/09/2010"),
                ("01253","COLGATE PALMOLIVE LTDA.","15/09/2010"),
                ("20074","COLGATE-PALMOLIVE CHILE S.A.","15/09/2010"),
                ("00843","COLGATE-PALMOLIVE COMPANY","15/09/2010"),
                ("36104","COLGATE-PALMOLIVE EUROPE SARL","29/12/2011"),
                ("31496","COLGATE-PALMOLIVE INVESTMENTS (BVI)","15/09/2010"),
                ("06361","COLGATE-PALMOLIVE S.p.A","15/09/2010"),
                ("08939","COMÉRCIO DE ESCOVA ORAL LTDA.","15/09/2010"),
                ("12362","DALTEX INDUSTRIAL LTDA","15/09/2010"),
                ("15394","FISBRA SERVIÇOS EM CONSULTORIA LTDA.","15/09/2010"),
                ("14292","HILL'S PET NUTRITION, INC.","15/09/2010"),
                ("11271","INDÚSTRIA E COMÉRCIO SANTA THEREZA LTDA.","15/09/2010"),
                ("05911","INDÚSTRIAS GESSY LEVER LTDA.","15/09/2010"),
                ("11664","KOLYNOS DO BRASIL LTDA","15/09/2010"),
                ("11124","ODOL SOCIEDAD ANONIMA INDUSTRIAL Y COMERCIAL","15/09/2010"),
                ("15471","RALSTON PURINA COMPANY","15/09/2010"),
                ("31602","TEC IMPORTS IMPORTAÇÃO E EXPORTAÇÃO LTDA.","15/09/2010"),
                ("07236","THE MENNEN COMPANY","15/09/2010"),
                ("12108","UNILEVER N.V.","15/09/2010"),
                ("08520","WYETH","15/09/2010"),
                ("00160","WYETH HOLDINGS CORPORATION","15/09/2010"),
                ("16118","ARCOM S/A.","15/09/2010"),
                ("22901","AROMATERAPIA INDÚSTRIA E COMÉRCIO LTDA.","15/09/2010"),
                ("09590","C-P INDÚSTRIA E COMÉRCIO LTDA.","15/09/2010"),
                ("01253","COLGATE PALMOLIVE LTDA.","15/09/2010"),
                ("20074","COLGATE-PALMOLIVE CHILE S.A.","15/09/2010"),
                ("00843","COLGATE-PALMOLIVE COMPANY","15/09/2010"),
                ("36104","COLGATE-PALMOLIVE EUROPE SARL","29/12/2011"),
                ("31496","COLGATE-PALMOLIVE INVESTMENTS (BVI)","15/09/2010"),
                ("06361","COLGATE-PALMOLIVE S.p.A","15/09/2010"),
                ("08939","COMÉRCIO DE ESCOVA ORAL LTDA.","15/09/2010"),
                ("12362","DALTEX INDUSTRIAL LTDA","15/09/2010"),
                ("15394","FISBRA SERVIÇOS EM CONSULTORIA LTDA.","15/09/2010"),
                ("14292","HILL'S PET NUTRITION, INC.","15/09/2010"),
                ("11271","INDÚSTRIA E COMÉRCIO SANTA THEREZA LTDA.","15/09/2010"),
                ("05911","INDÚSTRIAS GESSY LEVER LTDA.","15/09/2010"),
                ("11664","KOLYNOS DO BRASIL LTDA","15/09/2010"),
                ("11124","ODOL SOCIEDAD ANONIMA INDUSTRIAL Y COMERCIAL","15/09/2010"),
                ("15471","RALSTON PURINA COMPANY","15/09/2010"),
                ("31602","TEC IMPORTS IMPORTAÇÃO E EXPORTAÇÃO LTDA.","15/09/2010"),
                ("07236","THE MENNEN COMPANY","15/09/2010"),
                ("12108","UNILEVER N.V.","15/09/2010"),
                ("08520","WYETH","15/09/2010"),
                ("00160","WYETH HOLDINGS CORPORATION","15/09/2010"),
                ("16118","ARCOM S/A.","15/09/2010"),
                ("22901","AROMATERAPIA INDÚSTRIA E COMÉRCIO LTDA.","15/09/2010"),
                ("09590","C-P INDÚSTRIA E COMÉRCIO LTDA.","15/09/2010"),
                ("01253","COLGATE PALMOLIVE LTDA.","15/09/2010"),
                ("20074","COLGATE-PALMOLIVE CHILE S.A.","15/09/2010"),
                ("00843","COLGATE-PALMOLIVE COMPANY","15/09/2010"),
                ("36104","COLGATE-PALMOLIVE EUROPE SARL","29/12/2011"),
                ("31496","COLGATE-PALMOLIVE INVESTMENTS (BVI)","15/09/2010"),
                ("06361","COLGATE-PALMOLIVE S.p.A","15/09/2010"),
                ("08939","COMÉRCIO DE ESCOVA ORAL LTDA.","15/09/2010"),
                ("12362","DALTEX INDUSTRIAL LTDA","15/09/2010"),
                ("15394","FISBRA SERVIÇOS EM CONSULTORIA LTDA.","15/09/2010"),
                ("14292","HILL'S PET NUTRITION, INC.","15/09/2010"),
                ("11271","INDÚSTRIA E COMÉRCIO SANTA THEREZA LTDA.","15/09/2010"),
                ("05911","INDÚSTRIAS GESSY LEVER LTDA.","15/09/2010"),
                ("11664","KOLYNOS DO BRASIL LTDA","15/09/2010"),
                ("11124","ODOL SOCIEDAD ANONIMA INDUSTRIAL Y COMERCIAL","15/09/2010"),
                ("15471","RALSTON PURINA COMPANY","15/09/2010"),
                ("31602","TEC IMPORTS IMPORTAÇÃO E EXPORTAÇÃO LTDA.","15/09/2010"),
                ("07236","THE MENNEN COMPANY","15/09/2010"),
                ("12108","UNILEVER N.V.","15/09/2010"),
                ("08520","WYETH","15/09/2010"),
                ("00160","WYETH HOLDINGS CORPORATION","15/09/2010"),
    };

            await new BrowserFetcher().DownloadAsync();

            using (var browser = await Puppeteer.LaunchAsync(new LaunchOptions
            {
                Headless = true,
                Args = new[] { "--no-sandbox", "--disable-setuid-sandbox" }
            }))
            using (var page = await browser.NewPageAsync())
            {
                var htmlContent = $@"
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset='UTF-8'>
            <style>
                body {{
                    font-family: Arial, sans-serif;
                    margin: 0;
                    padding: 0;
                    font-size: 11px;
                    line-height: 1.2;
                }}
                .header-container {{
                    margin-bottom: 30px;
                }}
                .header-line {{
                    display: flex;
                    justify-content: space-between;
                    margin-bottom: 5px;
                }}
                .header-left {{
                    font-weight: bold;
                    font-size: 11px;
                }}
                .header-right {{
                    font-size: 11px;
                }}
                .pagination {{
                    text-align: right;
                    font-size: 10px;
                    margin-bottom: 15px;
                }}
                .title-box {{
                    border: 2px solid #000000;
                    background-color: #ffffff;
                    padding: 8px;
                    text-align: center;
                    margin: 15px 0;
                    font-size: 14px;
                    font-weight: bold;
                }}
                table {{
                    width: 100%;
                    border-collapse: collapse;
                    margin: 15px 0;
                    font-size: 10px;
                }}
                th, td {{
                    border: 1px solid #000000;
                    padding: 6px;
                    text-align: left;
                }}
                th {{
                    background-color: #f8f8f8;
                    font-weight: bold;
                    text-align: center;
                    font-size: 10px;
                }}
                td {{
                    font-size: 10px;
                }}
                .data-cell {{
                    text-align: right;
                    padding-right: 10px;
                }}
                .total-container {{
                    margin-top: 20px;
                    text-align: right;
                }}
                .total-box {{
                    border: 1px solid #000000;
                    background-color: #ffffff;
                    padding: 6px 12px;
                    display: inline-block;
                    font-weight: bold;
                    font-size: 10px;
                }}
                .page-break {{
                    page-break-after: always;
                }}
                @media print {{
                    .header-container {{
                        position: fixed;
                        top: 30px;
                        left: 40px;
                        right: 40px;
                    }}
                    .title-box {{
                        margin-top: 80px;
                    }}
                    body {{
                        margin-top: 100px;
                        margin-bottom: 50px;
                    }}
                    .total-container {{
                        position: fixed;
                        bottom: 30px;
                        right: 40px;
                    }}
                }}
            </style>
        </head>
        <body>
            <!-- Header para cada página -->
            <div class='header-container'>
                <div class='header-line'>
                    <div class='header-left'>{escritorio}</div>
                    <div class='header-right'>{dataHoje}</div>
                </div>
                <div class='pagination'>
                    Página <span class='pageNumber'></span> de <span class='totalPages'></span>
                </div>
            </div>

            <!-- Título do relatório -->
            <div class='title-box'>
                Grupo Econômico {titular}
            </div>

            <!-- Tabela de empresas -->
            <table>
                <thead>
                    <tr>
                        <th>Empresas Associadas</th>
                        <th>Data</th>
                    </tr>
                </thead>
                <tbody>";

                foreach (var e in empresas)
                {
                    htmlContent += $@"
                    <tr>
                        <td>{e.Codigo} - {e.Nome}</td>
                        <td class='data-cell'>{e.Data}</td>
                    </tr>";
                }

                htmlContent += $@"
                </tbody>
            </table>

            <!-- Total de empresas -->
            <div class='total-container'>
                <div class='total-box'>
                    Total: {empresas.Count}
                </div>
            </div>

            <script>
                // Atualizar numeração de páginas
                function updatePageNumbers() {{
                    var pageCount = Math.ceil(document.body.scrollHeight / 1056);
                    var pageNumbers = document.querySelectorAll('.pageNumber');
                    var totalPages = document.querySelectorAll('.totalPages');
                    
                    for (var i = 0; i < pageNumbers.length; i++) {{
                        pageNumbers[i].textContent = (i + 1);
                        totalPages[i].textContent = pageCount;
                    }}
                }}
                
                updatePageNumbers();
            </script>
        </body>
        </html>";

                await page.SetContentAsync(htmlContent);

                var pdfOptions = new PdfOptions
                {
                    Format = PaperFormat.A4,
                    MarginOptions = new MarginOptions
                    {
                        Top = "80px",
                        Right = "40px",
                        Bottom = "60px",
                        Left = "40px"
                    },
                    PrintBackground = true,
                    DisplayHeaderFooter = false
                };

                var pdfBytes = await page.PdfDataAsync(pdfOptions);
                return File(pdfBytes, "application/pdf", "relatorio-grupo-economico.pdf");
            }
        }

    }
}
