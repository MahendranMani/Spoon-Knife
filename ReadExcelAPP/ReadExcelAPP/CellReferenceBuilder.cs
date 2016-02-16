using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadExcelAPP
{
    internal interface ICellReferenceBuilder
    {
        string GetCellReference(int rowIndex, int columnIndex);
        string GetAbsoluteCellReference(int rowIndex, int columnIndex);
    }

    internal class CellReferenceBuilder : ICellReferenceBuilder
    {
        public string GetCellReference(int rowIndex, int columnIndex)
        {
            var displayRowIndex = rowIndex + 1;
            var row = displayRowIndex.ToString();
            var column = BuildColumnReference(columnIndex);
            var cellReference = column + row;
            return cellReference;
        }

        public string GetAbsoluteCellReference(int rowIndex, int columnIndex)
        {
            var displayRowIndex = rowIndex + 1;
            var row = displayRowIndex.ToString();
            var column = BuildColumnReference(columnIndex);
            var cellReference = string.Format("${0}${1}", column, row);
            return cellReference;
        }

        private string BuildColumnReference(int columnIndex)
        {
            var column = string.Empty;
            var remainder = columnIndex + 1;

            while (remainder > 0)
            {
                var i = (remainder - 1) % 26;
                column = (char)(65 + i) + column;
                remainder = (remainder - i) / 26;
            }

            return column;
        }
    }
}
