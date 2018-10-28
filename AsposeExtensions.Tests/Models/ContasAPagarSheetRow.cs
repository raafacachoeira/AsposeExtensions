using System;

namespace AsposeExtensions.Tests
{
    /// <summary>
    /// Classe de teste onde � usado as tr�s formas poss�veis de carregar os dados de uma coluna:
    /// 1 - usando construtor vazio e assim pegando pelo nome da propridade
    /// 2 - usando construtor passando o numero da posicao da coluna
    /// 3 - usando construtor passando o nome da coluna
    /// 4 - usando construtor passando o nome da coluna minusculo
    /// 5 - usando construtor passando o nome da coluna com espa�os
    /// Esta classe herda de SheetRowBase para poder carregar tamb�m a linha e nome da planilha
    /// </summary>
    public class ContasAPagarSheetRow : SheetRowBase
    {
        public ContasAPagarSheetRow(string sheetName, int row) 
            : base(sheetName, row)
        {
        }

        [SheetColumn]
        public DateTime? Data { get; set; }
        [SheetColumn(1)]
        public string Historico { get; set; }
        [SheetColumn("Doc.")]
        public int Doc { get; set; }
        [SheetColumn(3)]
        public decimal? Valor { get; set; }
        [SheetColumn("vcto.")]
        public DateTime DataDeVencimento { get; set; }
        [SheetColumn(5)]
        public string Pagto { get; set; }
        [SheetColumn(" Saldo ")]
        public string ValorFinal { get; set; }
    }
}
