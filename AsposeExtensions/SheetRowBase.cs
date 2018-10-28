namespace AsposeExtensions
{

    //TODO ter a interface abaixo para dar mais opção de herança e uso das propridades SheetName e Row. 
    // É necessário alterar o método RowToModel para ele não setar os valores dessas propriedades apenas via construtor.

    //public interface ISheetRowBase
    //{
    //    string SheetName { get; set; }
    //    int Row { get;  set; }
    //}

    public abstract class SheetRowBase
    {
        /// <summary>
        /// Classe base que terá a linha e o nome da planilha importada
        /// </summary>
        /// <param name="sheetName">Nome da panilha</param>
        /// <param name="row">Número da linha</param>
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
