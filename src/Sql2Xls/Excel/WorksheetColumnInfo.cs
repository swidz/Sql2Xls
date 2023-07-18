using Microsoft.Extensions.Logging;
using System.Data;

namespace Sql2Xls.Excel
{
    public class WorksheetColumnInfo
    {
        public bool IsInteger { get; private set; }
        public bool IsFloat { get; private set; }
        public bool IsDate { get; private set; }
        public string Code { get; private set; }
        public string ColumnName { get; private set; }
        public string Caption { get; private set; }
        public int Index { get; private set; }
        public bool IsInlineString { get; private set; }
        public bool IsSharedString { get; private set; }
        public Type DataType { get; private set; }
        public bool DateTimeAsString { get; private set; }

        private WorksheetColumnInfo()
        {
            DateTimeAsString = true;
        }

        public WorksheetColumnInfo(DataColumn dataColumn, int idx)
            : this()
        {
            Index = idx;
            ColumnName = dataColumn.ColumnName;
            Caption = dataColumn.Caption ?? dataColumn.ColumnName;
            Code = GetColumnName(idx);
            DataType = dataColumn.DataType;
            IsFloat = CheckIsFloat(dataColumn.DataType);
            IsDate = CheckIsDate(dataColumn.DataType);
            IsSharedString = CheckIsSharedString(dataColumn.DataType);
            IsInlineString = CheckIsInlineString(dataColumn.DataType);
            IsInteger = CheckIsInteger(dataColumn.DataType);

            DateTimeAsString = ExcelExportContext.Default.DateTimeAsString;
        }

        public WorksheetColumnInfo(DataColumn dataColumn, int idx, ExcelExportContext context)
            : this(dataColumn, idx)
        {
            DateTimeAsString = context.DateTimeAsString;
        }

        public WorksheetColumnInfo(IDataRecord record, int idx)
            : this()
        {
            Index = idx;
            ColumnName = record.GetName(idx);
            Caption = ColumnName;
            Code = GetColumnName(idx);
            DataType = record.GetFieldType(idx);
            IsFloat = CheckIsFloat(DataType);
            IsDate = CheckIsDate(DataType);
            IsSharedString = CheckIsSharedString(DataType);
            IsInlineString = CheckIsInlineString(DataType);
            IsInteger = CheckIsInteger(DataType);

            DateTimeAsString = ExcelExportContext.Default.DateTimeAsString;
        }

        public WorksheetColumnInfo(IDataRecord record, int idx, ExcelExportContext context)
            : this(record, idx)
        {
            DateTimeAsString = context.DateTimeAsString;
        }

        private bool CheckIsFloat(Type dataType)
        {
            if (dataType == typeof(decimal) || dataType == typeof(double) || dataType == typeof(float))
                return true;
            return false;
        }

        private bool CheckIsDate(Type dataType)
        {
            if (dataType == typeof(DateTime))
                return true;
            return false;
        }

        private bool CheckIsInteger(Type dataType)
        {
            if (dataType == typeof(short) || dataType == typeof(int) || dataType == typeof(long)
                || dataType == typeof(byte) || dataType == typeof(sbyte)
                || dataType == typeof(ushort) || dataType == typeof(uint) || dataType == typeof(ulong))
                return true;

            return false;
        }

        private bool CheckIsSharedString(Type dataType)
        {
            if (dataType == typeof(string) || dataType == typeof(char) || dataType == typeof(Guid))
                return true;

            if (dataType == typeof(DateTime) && DateTimeAsString == true)
                return true;

            return false;
        }

        private bool CheckIsInlineString(Type dataType)
        {
            return false;
        }

        private static readonly string[] columnCodesStatic = new string[] {
            "A", "B", "C", "D", "E", "F", "G", "H", "I", "J",
            "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T",
            "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD",
            "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN",
            "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX",
            "AY", "AZ", "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH",
            "BI", "BJ", "BK", "BL", "BM", "BN", "BO", "BP", "BQ", "BR",
            "BS", "BT", "BU", "BV", "BW", "BX", "BY", "BZ", "CA", "CB",
            "CC", "CD", "CE", "CF", "CG", "CH", "CI", "CJ", "CK", "CL",
            "CM", "CN", "CO", "CP", "CQ", "CR", "CS", "CT", "CU", "CV",
            "CW", "CX", "CY", "CZ", "DA", "DB", "DC", "DD", "DE", "DF",
            "DG", "DH", "DI", "DJ", "DK", "DL", "DM", "DN", "DO", "DP",
            "DQ", "DR", "DS", "DT", "DU", "DV", "DW", "DX", "DY", "DZ",
            "EA", "EB", "EC", "ED", "EE", "EF", "EG", "EH", "EI", "EJ",
            "EK", "EL", "EM", "EN", "EO", "EP", "EQ", "ER", "ES", "ET",
            "EU", "EV", "EW", "EX", "EY", "EZ", "FA", "FB", "FC", "FD",
            "FE", "FF", "FG", "FH", "FI", "FJ", "FK", "FL", "FM", "FN",
            "FO", "FP", "FQ", "FR", "FS", "FT", "FU", "FV", "FW", "FX",
            "FY", "FZ", "GA", "GB", "GC", "GD", "GE", "GF", "GG", "GH",
            "GI", "GJ", "GK", "GL", "GM", "GN", "GO", "GP", "GQ", "GR"
        };

        public static string GetColumnName(int index)
        {
            if (index < columnCodesStatic.Length)
                return columnCodesStatic[index];

            const byte BASE = 'Z' - 'A' + 1;
            string name = string.Empty;
            do
            {
                name = Convert.ToChar('A' + index % BASE) + name;
                index = index / BASE - 1;
            }
            while (index >= 0);
            return name;
        }
    }
}
