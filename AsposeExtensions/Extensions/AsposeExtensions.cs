using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Reflection;

namespace AsposeExtensions
{
    public static class AsposeExtensions
    {
        #region Import

        /// <summary>
        /// Importa uma unica linha para uma classe.
        /// </summary>
        /// <typeparam name="TTarget"></typeparam>
        /// <param name="excel">Instancia de Aspose Workbook</param>
        /// <param name="sheetName">Nome da planilha</param>
        /// <param name="row"></param>
        /// <returns></returns>
        public static TTarget RowToModel<TTarget>(this Workbook excel, string sheetName, int row, int rowHeader = 1)
        {
            return GetWorkSheetByName(excel, sheetName).RowToModel<TTarget>(row, rowHeader);
        }
        public static TTarget RowToModel<TTarget>(this Worksheet workSheet, int row, int rowHeader = 1)
        {
            var result = typeof(SheetRowBase).IsAssignableFrom(typeof(TTarget))
                ? (TTarget)Activator.CreateInstance(typeof(TTarget), workSheet.Name, row)
                : (TTarget)Activator.CreateInstance(typeof(TTarget));

            var properties = typeof(TTarget).GetProperties();

            foreach (var property in properties)
            {
                var attribute = property.GetCustomAttribute<SheetColumnAttribute>();

                if (attribute == null) continue;

                object cellValue;

                if (attribute.Column == null)
                {
                    var nameColumn = string.IsNullOrWhiteSpace(attribute.NameColumn)
                        ? property.Name
                        : attribute.NameColumn;

                    cellValue = workSheet.Cells.GetCellValueByNameColumn(row - 1, nameColumn, rowHeader);
                }
                else
                {
                    cellValue = workSheet.Cells.GetCellValueByPositionColumn(row - 1, attribute.Column.Value);
                }

                var typeTarget = Nullable.GetUnderlyingType(property.PropertyType) ?? property.PropertyType;
                var cellValueConvert = cellValue == null
                                                ? null
                                                : cellValue.GetType().Equals(typeTarget)
                                                    ? cellValue
                                                    : Convert.ChangeType(cellValue, typeTarget, null);

                property.SetValue(result, cellValueConvert);

            }

            return result;
        }

        public static IList<TTarget> RowsToModelList<TTarget>(this Workbook excel, string sheetName, int rowHeader = 1)
        {
            return GetWorkSheetByName(excel, sheetName).RowsToModelList<TTarget>(rowHeader);
        }
        public static IList<TTarget> RowsToModelList<TTarget>(this Worksheet workSheet, int rowHeader = 1)
        {
            IList<TTarget> result = (List<TTarget>)Activator.CreateInstance(typeof(List<TTarget>));

            for (int i = rowHeader; i < workSheet.Cells.Rows.Count; i++)
            {
                var rowModel = workSheet.RowToModel<TTarget>(i + 1, rowHeader);
                result.Add(rowModel);
            }

            return result;
        }

        private static Worksheet GetWorkSheetByName(Workbook excel, string sheetName)
        {
            if (string.IsNullOrEmpty(sheetName))
                throw new ArgumentNullException(nameof(sheetName));

            var workSheet = excel.Worksheets[sheetName];
            if (workSheet == null)
                throw new AsposeExtensionsSheetNotFoundException(sheetName);

            workSheet.Cells.DeleteBlankRows();

            return workSheet;
        }


        public static object GetCellValueByPositionColumn(this Cells cells, int row, int positionColumn)
        {
            return cells[row, positionColumn].Value;
        }
        public static object GetCellValueByNameColumn(this Cells cells, int row, string nameColumn, int rowHeader = 1)
        {
            var enumerator = cells.GetRow(rowHeader - 1).GetEnumerator();
            while (enumerator.MoveNext())
            {
                var cell = enumerator.Current as Cell;
                if (!cell.Type.Equals(CellValueType.IsNull)
                    && cell.Value.ToString().Trim().ToLower().Equals(nameColumn.Trim().ToLower()))
                {
                    return cells[row, cell.Column].Value;
                }
            }

            return null;
        }

        #endregion

        #region Export

        //todo

        #endregion

    }
}
