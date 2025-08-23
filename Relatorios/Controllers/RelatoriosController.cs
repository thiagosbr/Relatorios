using Microsoft.AspNetCore.Mvc;
using PuppeteerSharp;
using PuppeteerSharp.Media;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using System.Data.Common;
using System.Text;

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


        [HttpGet("grupo-economico-puppeteer-HTML")]
        public async Task<IActionResult> GrupoEconomicoPuppeteerHTML()
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
            ("00160","WYETH HOLDINGS CORPORATION","15/09/2010")
            // ... (mantenha sua lista completa)
        };

            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions
            {
                Headless = true,
                Args = new[] { "--no-sandbox", "--disable-setuid-sandbox" }
            });
            await using var page = await browser.NewPageAsync();

            // -------- HTML + CSS que replica o PDF --------
            var sb = new StringBuilder();
            sb.Append($@"
                <!DOCTYPE html>
                <html lang='pt-BR'>
                <head>
                <meta charset='utf-8'>
                <style>
                  /* Página A4 controlada pelo Puppeteer via PdfOptions */
                  html, body {{
                    margin: 0;
                    padding: 0;
                    -webkit-print-color-adjust: exact;
                    print-color-adjust: exact;
                    font-family: Arial, Helvetica, sans-serif;
                    font-size: 11px;
                    line-height: 1.25;
                    color: #000;
                  }}

                  /* Tabela principal: usamos THEAD para repetir título e cabeçalho em cada página */
                  table.grid {{
                    width: 100%;
                    border-collapse: collapse;
                  }}

                  thead {{ display: table-header-group; }}
                  tfoot {{ display: table-footer-group; }} /* caso precise no futuro */
                  tr {{ page-break-inside: avoid; }}

                  /* Título que repete em todas as páginas (fica dentro do THEAD) */
                  .title-cell {{
                    border: 2px solid #000;
                    padding: 8px 6px;
                    text-align: center;
                    font-weight: 700;
                    font-size: 12px;
                    background: #fff;
                  }}

                  /* Cabeçalho da tabela */
                  .header-th {{
                    border: 1px solid #000;
                    padding: 6px;
                    text-align: center;
                    font-weight: 700;
                    background: #f5f5f5;
                    font-size: 10px;
                  }}

                  /* Células do corpo */
                  td {{
                    border: 1px solid #000;
                    padding: 6px;
                    vertical-align: top;
                    font-size: 10px;
                  }}
                  .date {{ text-align: right; }}

                  /* Total ao final do documento */
                  .total-box {{
                            margin-top: 20px;
                            border: 1px solid #000;
                            padding: 6px 0;
                            font-weight: 700;
                            font-size: 10px;
                            text-align: center;   /* centraliza o texto */
                            width: 100%;          /* ocupa toda a largura */
                            box-sizing: border-box;
                  }}
                </style>
                </head>
                <body>

                  <table class='grid'>
                    <thead>
                      <tr>
                        <th class='title-cell' colspan='2'>Grupo Econômico {titular}</th>
                      </tr>
                      <tr>
                        <th class='header-th'>Empresas Associadas</th>
                        <th class='header-th'>Data</th>
                      </tr>
                    </thead>
                    <tbody>");
                            foreach (var e in empresas)
                                sb.Append($@"
                      <tr>
                        <td>{e.Codigo} - {e.Nome}</td>
                        <td class='date'>{e.Data}</td>
                      </tr>");
                            sb.Append($@"
                    </tbody>
                  </table>

                  <div class='total-box'>Total: {empresas.Count}</div>

                </body>
                </html>");

                            var html = sb.ToString();

                            // Header do PDF (repetido em todas as páginas) – usa placeholders do Chromium
                            var headerTemplate = $@"
                <div style='width:100%; padding:0 40px; box-sizing:border-box; font-family:Arial,Helvetica,sans-serif; color:#000;'>
                  <div style='display:flex; justify-content:space-between; font-size:11px;'>
                    <span style='font-weight:700;'>{escritorio}</span>
                    <span>{dataHoje}</span>
                  </div>
                  <div style='text-align:right; font-size:10px; margin-top:2px;'>
                    Página <span class='pageNumber'></span> de <span class='totalPages'></span>
                  </div>
                </div>";

            // Rodapé vazio (apenas reserva espaço para bottom margin)
            var footerTemplate = @"<div style='width:100%; height:0;'></div>";

            await page.SetContentAsync(html);

            var pdfOptions = new PdfOptions
            {
                Format = PaperFormat.A4,
                PrintBackground = true,
                DisplayHeaderFooter = true,
                HeaderTemplate = headerTemplate,
                FooterTemplate = footerTemplate,
                MarginOptions = new MarginOptions
                {
                    // espaço para header/rodapé (ajuste fino se quiser)
                    Top = "90px",      // cabeçalho: 2 linhas
                    Bottom = "40px",
                    Left = "40px",
                    Right = "40px"
                },
                PreferCSSPageSize = false
            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);
            return File(pdfBytes, "application/pdf", "relatorio-grupo-economico-HTML.pdf");
        }

        [HttpGet("relatorio-pastas-marcas")]
        public IActionResult RelatorioPastasMarcas()
        {
            QuestPDF.Settings.License = LicenseType.Community;
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
                               .Text($"RELATÓRIO DE PASTA DE MARCAS")
                               .FontSize(12)
                               .Bold();

                            tituloCol.Item().PaddingTop(12);
                        });
                    });

                    page.Content().Column(col =>
                    {
                        col.Item().ShowOnce().Text("Não foi utilizado nenhum critério como filtro").Italic();
                        col.Item().ShowOnce().PageBreak();
                        HeaderLinha(col, "933739230 - LORD & BERRY", "M301089");

                        col.Item().PaddingVertical(6);

                        col.Item().Border(1).Padding(8).Row(r =>
                        {
                            r.ConstantItem(120).Border(1).Height(80)
                                .AlignCenter().AlignMiddle().Image("Imagem.jpg");

                            r.RelativeItem().PaddingLeft(8).Column(c =>
                            {
                                AddLabelValue(c, "Status", "Pedido");
                                AddLabelValue(c, "Situação", "Em vigor");
                                AddLabelValue(c, "Apresentação", "Nominativa");
                                AddLabelValue(c, "Setor", "MARCAS");
                                AddLabelValue(c, "Natureza", "De Produto");
                                AddLabelValue(c, "IRN", "");
                                AddLabelValue(c, "Cliente", "METROCONSULT S.R.L.");
                                AddLabelValue(c, "Cliente de Cobrança", "METROCONSULT SRL");
                                AddLabelValue(c, "Titular(es)", "LORD & BERRY EUROPE S.R.L.");
                            });

                            r.ConstantItem(170).PaddingLeft(10).Column(c =>
                            {
                                AddLabelValue(c, "Data Dep.", "05/03/2024");
                                AddLabelValue(c, "Data 1º Reg", "");
                                AddLabelValue(c, "Último Reg", "");
                                AddLabelValue(c, "Prorrogação", "");
                            });
                        });

                        col.Item().PaddingVertical(8);

                        col.Item().Border(2).BorderColor(Colors.Black)
                               .Background(Colors.White)
                               .Padding(1)
                               .AlignCenter().Text("NCL");
                        col.Item().Border(1).Padding(8).Column(ncl =>
                        {
                            ncl.Item().Row(r =>
                            {
                                r.ConstantItem(50).Text("Código").SemiBold();
                                r.ConstantItem(40).Text("03");
                                r.ConstantItem(60).Text("Países").SemiBold();
                                r.RelativeItem().Text("Brasil");
                            });

                            ncl.Item().PaddingTop(6).Text("Produtos e serviços (PT)").Bold();
                            ncl.Item().Text("Cosméticos; maquiagem para os olhos; batons; lápis labiais; bálsamo labial; brilho labial; óleo labial; corretivo; base em creme; pó facial.");

                            ncl.Item().PaddingTop(6).Text("Produtos e serviços (EN)").Bold();
                            ncl.Item().Text("Cosmetics; eye makeup; lipsticks; lip pencils; lip balm; lip gloss; lip oil; concealer; cream foundation; face powder.");
                        });

                        col.Item().PaddingVertical(8);

                        col.Item().Border(2).BorderColor(Colors.Black)
                             .Background(Colors.White)
                             .Padding(1)
                             .AlignCenter().Text("RPI");
                        col.Item().Border(1).Padding(8)
                            .Text("Nº2778 - 02/04/2024 - IPAS0090000 - Publicação de pedido de registro para oposição");
                        col.Item().PaddingVertical(12);

                        // ===== Marca (Registro) com RPI =====
                        HeaderLinha(col, "750067527 - ARIOLI", "M301775");
                        col.Item().PaddingVertical(6);

                        col.Item().Border(1).Padding(8).Row(r =>
                        {
                            r.ConstantItem(120).Border(1).Height(80)
                                .AlignCenter().AlignMiddle().Text("Mista").Bold();

                            r.RelativeItem().PaddingLeft(8).Column(c =>
                            {
                                AddLabelValue(c, "Status", "Registro");
                                AddLabelValue(c, "Situação", "Em vigor");
                                AddLabelValue(c, "Apresentação", "Mista");
                                AddLabelValue(c, "Setor", "MARCAS");
                                AddLabelValue(c, "Natureza", "De Produto");
                                AddLabelValue(c, "IRN", "");
                                AddLabelValue(c, "Cliente", "METROCONSULT S.R.L.");
                                AddLabelValue(c, "Cliente de Cobrança", "METROCONSULT S.R.L.");
                                AddLabelValue(c, "Titular(es)", "ARIOLI S.P.A.");
                            });

                            r.ConstantItem(170).PaddingLeft(10).Column(c =>
                            {
                                AddLabelValue(c, "Data Dep.", "22/04/1975");
                                AddLabelValue(c, "Data 1º Reg", "16/02/1982");
                                AddLabelValue(c, "Último Reg", "");
                                AddLabelValue(c, "Prorrogação", "16/02/2032");
                            });
                        });

                        col.Item().PaddingVertical(8);
                        col.Item().Border(2).BorderColor(Colors.Black)
                            .Background(Colors.White)
                            .Padding(1)
                            .AlignCenter().Text("RPI");
                        col.Item().Border(1).Padding(8).Column(rpi =>
                        {
                            rpi.Item().Text("Nº1095 - 26/11/1991 - 990 - CONCEDIDA PRORROGAÇÃO");
                            rpi.Item().Text("Nº1663 - 19/11/2002 - 565 - ANOTADA TRANSFERÊNCIA");
                            rpi.Item().Text("Nº1814 - 11/10/2005 - 990 - CONCEDIDA PRORROGAÇÃO");
                            rpi.Item().Text("Nº2283 - 07/10/2014 - IPAS2703741 - Prorrogação de registro de marca e expedição de certificado no prazo ordinário (374.1)");
                            rpi.Item().Text("Nº2353 - 10/02/2016 - IPAS5778243 - Emissão de folha de rosto de cópia reprográfica simples");
                            rpi.Item().Text("Nº2473 - 29/05/2018 - IPAS2673483 - Exigência de mérito");
                            rpi.Item().Text("Nº2487 - 04/09/2018 - IPAS2703483 - Deferimento de petição");
                            rpi.Item().Text("Nº2668 - 22/02/2022 - IPAS2703745 - Deferimento de petição");
                            rpi.Item().Text("Nº2746 - 22/08/2023 - IPAS5663663 - Petição de retificação atendida");
                        });
                        col.Item().PaddingVertical(12);

                        // ===== Transferência =====
                        col.Item().Text("Transferência").Bold().FontSize(12).Italic() ;

                        col.Item().Border(1).Padding(8).Column(tr =>
                        {
                            tr.Item().Row(r =>
                            {
                                r.RelativeItem().Text(t =>
                                {
                                    t.Span("Pasta ").SemiBold(); t.Span("TM220263");
                                });
                                r.RelativeItem().Text(t =>
                                {
                                    t.Span("Setor ").SemiBold(); t.Span("MARCAS");
                                });
                            });

                            tr.Item().Row(r =>
                            {
                                r.RelativeItem().Text(t => { t.Span("Situação ").SemiBold(); t.Span("Em vigor"); });
                                r.RelativeItem().Text(t => { t.Span("Responsável ").SemiBold(); t.Span("AMCendon"); });
                            });

                            tr.Item().Row(r =>
                            {
                                r.RelativeItem().Text(t => { t.Span("Data de Depósito ").SemiBold(); t.Span("26/06/2025"); });
                                r.RelativeItem().Text(t => { t.Span("Protocolo ").SemiBold(); t.Span("850250334306"); });
                            });

                            tr.Item().Row(r =>
                            {
                                r.RelativeItem().Text(t => { t.Span("Usuário ").SemiBold(); t.Span("Ana Maria Guimaraes Cendon"); });
                                r.RelativeItem().Text(t => { t.Span("Contato / E-mail ").SemiBold(); t.Span(""); });
                            });

                            tr.Item().PaddingTop(6).Row(r =>
                            {
                                r.RelativeItem().Text(t => { t.Span("Cliente ").SemiBold(); t.Span("Metroconsult S.r.l."); });
                                r.RelativeItem().Text(t => { t.Span("Cliente de nota ").SemiBold(); t.Span("Metroconsult S.r.l."); });
                            });

                            tr.Item().PaddingTop(6).Text(t => { t.Span("Titular Novo ").SemiBold(); t.Span("INVESTIMENTI INDUSTRIALI ITALIA 01 S.R.L."); });
                            tr.Item().Text(t => { t.Span("Titular Antigo ").SemiBold(); t.Span("ARIOLI S.P.A."); });

                            tr.Item().PaddingTop(8).Text("Histórico").Bold();
                            tr.Item().Table(table =>
                            {
                                table.ColumnsDefinition(c =>
                                {
                                    c.ConstantColumn(40);
                                    c.RelativeColumn();
                                    c.RelativeColumn();
                                });

                                table.Header(h =>
                                {
                                    h.Cell().Border(1).Padding(4).Text("Tipo").SemiBold();
                                    h.Cell().Border(1).Padding(4).Text("De").SemiBold();
                                    h.Cell().Border(1).Padding(4).Text("Para").SemiBold();
                                });

                                table.Cell().Border(1).Padding(4).Text("TF");
                                table.Cell().Border(1).Padding(4).Text("ARIOLI S.P.A.");
                                table.Cell().Border(1).Padding(4).Text("INVESTIMENTI INDUSTRIALI ITALIA 01 S.R.L.");
                            });
                        });

                        col.Item().PaddingVertical(12);

                        // ===== Oposição / Marca de Terceiro =====
                        col.Item().Text("Marca de Terceiro").Bold().FontSize(12).Italic();
                        HeaderLinha(col, "931476054 - GEOLOGAR", "O210236");

                        col.Item().PaddingVertical(6);

                        col.Item().Border(1).Padding(8).Row(r =>
                        {
                            r.ConstantItem(120).Border(1).Height(80)
                                .AlignCenter().AlignMiddle().Text("Mista").Bold();

                            r.RelativeItem().PaddingLeft(8).Column(c =>
                            {
                                AddLabelValue(c, "Status", "");
                                AddLabelValue(c, "Situação", "Em vigor");
                                AddLabelValue(c, "Apresentação", "Mista");
                                AddLabelValue(c, "Setor", "MARCAS");
                                AddLabelValue(c, "Natureza", "Produtos/Serviços");
                                AddLabelValue(c, "IRN", "");
                                AddLabelValue(c, "Cliente", "METROCONSULT S.R.L.");
                                AddLabelValue(c, "Cliente de Cobrança", "METROCONSULT S.R.L.");
                                AddLabelValue(c, "Titular(es)", "GEOLOG INTERNATIONAL B.V.");
                            });

                            r.ConstantItem(170).PaddingLeft(10).Column(c =>
                            {
                                AddLabelValue(c, "Data Dep.", "09/08/2023");
                                AddLabelValue(c, "Data 1º Reg", "");
                                AddLabelValue(c, "Prorrogação", "");
                            });
                        });

                        col.Item().PaddingVertical(8);

                        col.Item().Border(2).BorderColor(Colors.Black)
                          .Background(Colors.White)
                          .Padding(1)
                          .AlignCenter().Text("RPI");
                        col.Item().Border(1).Padding(8).Column(rpi =>
                        {
                            
                            rpi.Item().Text("Nº2761 - 05/12/2023 - IPAS4230000 -");
                            rpi.Item().Text("Nº2748 - 05/09/2023 - IPAS0090000 -");
                        });

                        col.Item().PaddingTop(12).Border(2).BorderColor(Colors.Black)
                          .Padding(1)
                          .AlignCenter().Text(t => { t.Span("Total de Pastas: ").SemiBold(); t.Span("4"); });


                    });
                });
            }).GeneratePdf();

            return File(pdf, "application/pdf", "relatorio-Marcas.pdf");
        }

        
        private static void AddLabelValue(ColumnDescriptor col, string label, string value)
        {
            col.Item().Text(t =>
            {
                t.Span(label + " ").SemiBold();
                t.Span(value ?? string.Empty);
            });
        }

        private static void HeaderLinha(ColumnDescriptor col, string esquerda, string direita)
        {
            col.Item().Row(r =>
            {
                r.RelativeItem().Text(esquerda).Bold().FontSize(12);
                r.ConstantItem(100).AlignRight().Text(direita).Bold().FontSize(12);
            });
        }

    }

 
}

