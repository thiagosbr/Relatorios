using Microsoft.AspNetCore.Mvc;
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

    }
}
