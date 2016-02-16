using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadExcelAPP
{
    public enum CellReadStatus
    {
        Success,
        EmptyCell,
        WrongDataType
    }

    public class CellReadResponse<T>
    {
        public CellReadStatus Status { get; set; }
        public T Value { get; set; }
    }


}
