using System.Diagnostics;
using System.Reflection;
using System.Text.Json;
using GeradorRelatorioPDF;
using iTextSharp.text;
using iTextSharp.text.pdf;

public class Program
{
    static List<Pessoa> pessoas = [];
    static readonly string exedir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
    static readonly BaseFont fonteBase = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, false);

    private static void Main(string[] args)
    {
        DesserealizarPessoas();
        GerarRelatorioEmPDF(100);
    }
    static void DesserealizarPessoas()
    {
        string fileName = "pessoas.json";
        string path = Path.Combine(exedir, "Source", fileName);


        if (File.Exists(path))
        {

            using var sr = new StreamReader(path);
            var dados = sr.ReadToEnd();

            pessoas = JsonSerializer.Deserialize(dados, typeof(List<Pessoa>)) as List<Pessoa>;
        }
        else
        {
            Console.WriteLine("Arquivo não encontrado!");
        }
    }

    static void GerarRelatorioEmPDF(int qtdePessoas)
    {
        var pessoasSelecionadas = pessoas.Take(qtdePessoas).ToList();
        if (pessoasSelecionadas.Count >= 0)
        {
            //cálculo da quantidade total de páginas
            int totalPaginas = 1;
            int totalLinhas = pessoasSelecionadas.Count;
            if (totalLinhas > 24)
                totalPaginas += (int)Math.Ceiling((totalLinhas - 24) / 29F);

            //configuração do documento PDF
            var pxPorMm = 72 / 25.2F;
            var pdf = new iTextSharp.text.Document(PageSize.A4, 15 * pxPorMm, 15 * pxPorMm, 15 * pxPorMm, 20 * pxPorMm);
            var nomeArquivo = $"pessoas.{DateTime.Now.ToString("yyyy-MM-dd.HH.mm.ss")}.pdf";
            var pathWrite = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, nomeArquivo);
            var arquivo = new FileStream(pathWrite, FileMode.Create);
            var writer = PdfWriter.GetInstance(pdf, arquivo);
            writer.PageEvent = new EventosDePagina(totalPaginas);
            pdf.Open();

            //adição de título
            var fonteParagrafo = new iTextSharp.text.Font(fonteBase, 32, iTextSharp.text.Font.NORMAL, BaseColor.Black);
            var titulo = new Paragraph("Relatório de pessoas\n\n", fonteParagrafo)
            {
                Alignment = Element.ALIGN_LEFT,
                SpacingAfter = 4
            };
            pdf.Add(titulo);

            //adição da imagem
            var caminhoImagem = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "source//img", "youtube.png");
            if (File.Exists(caminhoImagem))
            {
                iTextSharp.text.Image logo = iTextSharp.text.Image.GetInstance(caminhoImagem);
                float razaoAlturaLargura = logo.Width / logo.Height;
                float alturaLogo = 32;
                float larguraLogo = alturaLogo * razaoAlturaLargura;
                logo.ScaleToFit(larguraLogo, alturaLogo);

                var margemEsquerda = pdf.PageSize.Width - pdf.RightMargin - larguraLogo;
                var margemTopo = pdf.PageSize.Height - pdf.TopMargin - 54;
                logo.SetAbsolutePosition(margemEsquerda, margemTopo);
                writer.DirectContent.AddImage(logo, false);
            }

            //adição de um link
            var fonteLink = new iTextSharp.text.Font(fonteBase, 9.9F, Font.NORMAL, BaseColor.Blue);
            var link = new Chunk("Canal do Prof. Ricardo Maroquio", fonteLink);
            // link.SetAnchor("https://www.youtube.com/maroquio");
            link.SetAnchor("https://www.youtube.com/watch?v=Gm2pJfCJyUw&t=887s");
            var larguraTextoLink = fonteBase.GetWidthPoint(link.Content, fonteLink.Size);
            var caixaTexto = new ColumnText(writer.DirectContent);
            caixaTexto.AddElement(link);
            caixaTexto.SetSimpleColumn(
                pdf.PageSize.Width - pdf.RightMargin - larguraTextoLink,
                pdf.PageSize.Height - pdf.TopMargin - (30 * pxPorMm),
                pdf.PageSize.Width - pdf.RightMargin,
                pdf.PageSize.Height - pdf.TopMargin - (18 * pxPorMm));
            caixaTexto.Go();

            //adição da tabela de dados
            var tabela = new PdfPTable(5);
            float[] larguraColunas = [0.6f, 2F, 1.5f, 1f, 1f];
            tabela.SetWidths(larguraColunas);
            tabela.DefaultCell.BorderWidth = 0;
            tabela.WidthPercentage = 100;

            //adição de células de títulos das colunas
            CriarCelulaTexto(tabela, "Código", PdfCell.ALIGN_CENTER, true);
            CriarCelulaTexto(tabela, "Nome", PdfCell.ALIGN_LEFT, true);
            CriarCelulaTexto(tabela, "Profissão", PdfCell.ALIGN_CENTER, true);
            CriarCelulaTexto(tabela, "Salário", PdfCell.ALIGN_CENTER, true);
            CriarCelulaTexto(tabela, "Empregada", PdfCell.ALIGN_CENTER, true);

            foreach (var p in pessoasSelecionadas)
            {
                CriarCelulaTexto(tabela, p.IdPessoa.ToString("D6"), PdfCell.ALIGN_CENTER, false, false, 10);
                CriarCelulaTexto(tabela, p.Nome + " " + p.Sobrenome, 0, false, false, 10);
                CriarCelulaTexto(tabela, p.Profissao.Nome, PdfCell.ALIGN_CENTER, false, false, 10);
                CriarCelulaTexto(tabela, p.Salario.ToString("#,##0.00"), PdfCell.ALIGN_RIGHT, false, false, 10);
                // CriarCelulaTexto(tabela, p.Empregado ? "Sim" : "Não", PdfCell.ALIGN_CENTER);

                var caminhoImagemCelula = p.Empregado ? "emoji_feliz.png" : "emoji_triste.png";
                caminhoImagem = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "source//img", caminhoImagemCelula);
                CriarCelulaImagem(tabela, caminhoImagem, 20, 20);
            }

            pdf.Add(tabela);

            pdf.Close();
            arquivo.Close();

            //abre o pdf no visualizador padrão;

            if (File.Exists(pathWrite))
            {
                Process.Start("open", pathWrite);
                // Process.Start(new ProcessStartInfo()
                // {
                //     Arguments = $"/c start {pathWrite}",
                //     FileName = "cmd.exe",
                //     CreateNoWindow = true
                // });
            }

        }
    }

    private static void CriarCelulaTexto(PdfPTable tabela, string texto, int alinhamentoHorz = PdfPCell.ALIGN_LEFT,
    bool negrito = false, bool italico = false, int tamanhoFonte = 12, int alturaCelula = 25)
    {
        int estilo = iTextSharp.text.Font.NORMAL;

        if (negrito && italico)
        {
            estilo = iTextSharp.text.Font.BOLDITALIC;
        }
        else if (negrito)
        {
            estilo = iTextSharp.text.Font.BOLD;
        }
        else if (italico)
        {
            estilo = iTextSharp.text.Font.ITALIC;
        }

        var fonteCelula = new iTextSharp.text.Font(fonteBase, tamanhoFonte, estilo, BaseColor.Black);

        var bgColor = new iTextSharp.text.BaseColor(1F, 1F, 1F);
        if (tabela.Rows.Count % 2 == 1)
            bgColor = new BaseColor(0.90F, 0.90F, 0.90F); // o F (0.90F) indica ponto flutuante, neste caso vai de 0 a 1

        var celula = new PdfPCell(new Phrase(texto, fonteCelula))
        {
            HorizontalAlignment = alinhamentoHorz,
            VerticalAlignment = PdfPCell.ALIGN_MIDDLE,
            Border = 0,
            BorderWidthBottom = 1,
            FixedHeight = alturaCelula,
            PaddingBottom = 5,
            BackgroundColor = bgColor
        };
        tabela.AddCell(celula);
    }

    static void CriarCelulaImagem(PdfPTable tabela, string caminhoImagem, int larguraImagem, int alturaImagem, int alturaCelula = 25)
    {
        var bgColor = new iTextSharp.text.BaseColor(1F, 1F, 1F);
        if (tabela.Rows.Count % 2 == 1)
            bgColor = new BaseColor(0.90F, 0.90F, 0.90F); // o F (0.90F) indica ponto flutuante, neste caso vai de 0 a 1


        if (File.Exists(caminhoImagem))
        {
            iTextSharp.text.Image imagem = iTextSharp.text.Image.GetInstance(caminhoImagem);
            imagem.ScaleToFit(larguraImagem, alturaImagem);

            var celula = new PdfPCell(imagem)
            {
                HorizontalAlignment = PdfPCell.ALIGN_CENTER,
                VerticalAlignment = PdfPCell.ALIGN_MIDDLE,
                Border = 0,
                BorderWidthBottom = 1,
                FixedHeight = alturaCelula,
                BackgroundColor = bgColor
            };
            tabela.AddCell(celula);
        }
    }
}