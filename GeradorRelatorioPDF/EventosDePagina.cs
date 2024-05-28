using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace GeradorRelatorioPDF
{
    public class EventosDePagina : PdfPageEventHelper
    {
        private PdfContentByte Wdc;
        private BaseFont FonteBaseRodape { get; set; }
        private iTextSharp.text.Font FonteRodape { get; set; }
        public int TotalPaginas { get; set; } = 1;

        public EventosDePagina(int totalPaginas)
        {
            FonteBaseRodape = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, false);
            FonteRodape = new iTextSharp.text.Font(FonteBaseRodape, 8f, iTextSharp.text.Font.NORMAL, BaseColor.Black);
            this.TotalPaginas = totalPaginas;
        }

        public override void OnEndPage(PdfWriter writer, Document document)
        {
            base.OnEndPage(writer, document);
            AdcionarMomentoGeracaoRelatorio(writer, document);
            AdicionarNumeroDasPaginas(writer, document);
        }

        public override void OnOpenDocument(PdfWriter writer, Document document)
        {
            base.OnOpenDocument(writer, document);
            this.Wdc = writer.DirectContent;
        }

        private void AdcionarMomentoGeracaoRelatorio(PdfWriter writer, Document document)
        {
            var textoMomentoGeracao = $"Gerado em {DateTime.Now.ToShortDateString()} às {DateTime.Now.ToShortTimeString()}";
            Wdc.BeginText();
            Wdc.SetFontAndSize(FonteRodape.BaseFont, FonteRodape.Size);
            Wdc.SetTextMatrix(document.LeftMargin, document.BottomMargin * 0.75f);
            Wdc.ShowText(textoMomentoGeracao);
            Wdc.EndText();
        }

        private void AdicionarNumeroDasPaginas(PdfWriter writer, Document document)
        {
            int paginaAtual = writer.PageNumber;
            var textoPaginacao = $"Página {paginaAtual} de {TotalPaginas}";

            float larguraTotalPaginacao = FonteBaseRodape.GetWidthPoint(textoPaginacao, FonteRodape.Size);

            var tamanhoPagina = document.PageSize;

            Wdc.BeginText();
            Wdc.SetFontAndSize(FonteRodape.BaseFont, FonteRodape.Size);
            Wdc.SetTextMatrix(tamanhoPagina.Width - document.RightMargin - larguraTotalPaginacao, document.BottomMargin * 0.75f);
            Wdc.ShowText(textoPaginacao);
            Wdc.EndText();
        }
    }
}