namespace AsposeExtensions
{

    //TODO ter a interface abaixo para dar mais op��o de heran�a e uso das propridades SheetName e Row. 
    // � necess�rio alterar o m�todo RowToModel para ele n�o setar os valores dessas propriedades apenas via construtor.

    //public interface ISheetRowBase
    //{
    //    string SheetName { get; set; }
    //    int Row { get;  set; }
    //}

    public abstract class SheetRowBase
    {
        /// <summary>
        /// Classe base que ter� a linha e o nome da planilha importada
        /// </summary>
        /// <param name="sheetName">Nome da panilha</param>
        /// <param name="row">N�mero da linha</param>
        protected SheetRowBase(string sheetName, int row)
        {
            SheetName = sheetName;
            Row = row;
        }

        public string SheetName { get; private set; }
        public int Row { get; private set; }

        public override string ToString()
        {
            return string.Format("Planilha: '{0}' | Linha: {1}", SheetName, Row);
        }
    }
}
