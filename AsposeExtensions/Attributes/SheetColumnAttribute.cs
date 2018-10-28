using System;

namespace AsposeExtensions
{
    [AttributeUsage(AttributeTargets.Property)]
    public class SheetColumnAttribute : Attribute
    {
        /// <summary>
        /// Se deseja buscar o valor pelo nome da propiedade da classe o construtor deve ser vazio
        /// </summary>
        public SheetColumnAttribute()
        {
        }

        /// <summary>
        /// Se deseja buscar o valor pelo nome da coluna
        /// </summary>
        /// <param name="nameColumn">Nome da coluna. Se passado null, irá buscar pelo nome da propiedade da classe</param>
        public SheetColumnAttribute(string nameColumn)
        {
            NameColumn = nameColumn;
        }

        /// <summary>
        /// Se deseja buscar o valor pela posicao da coluna
        /// </summary>
        /// <param name="positionColumn">Numero da coluna, iniciando em 0</param>
        public SheetColumnAttribute(int positionColumn)
        {
            Column = positionColumn;
        }

        /// <summary>
        /// Pode conter espaços
        /// </summary>
        public string NameColumn { get; private set; }

        /// <summary>
        /// Inicia em 0
        /// </summary>
        public int? Column { get; private set; }
    }
}
