using System;

namespace AsposeExtensions
{
    /// <summary>
    /// Thrown quando não encontrado o nome da planilha informada 
    /// </summary>
    [Serializable]
    public class AsposeExtensionsSheetNotFoundException : Exception
    {
        public AsposeExtensionsSheetNotFoundException(string sheetName) 
            : base(string.Format("A sheet with name \"{0}\" not found in excel file.", sheetName))
        {
        }
    }
}
