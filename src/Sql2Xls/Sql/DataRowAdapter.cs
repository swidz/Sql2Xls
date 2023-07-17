using System;
using System.Data;

namespace Sql2Xls.Sql;

public class DataRowAdapter : IDataRecord
{
    public DataRow Row { get; private set; }

    public DataRowAdapter(DataRow row)
    {
        Row = row;
    }

    public object this[string name]
    {
        get { return Row[name]; }
    }

    public object this[int i]
    {
        get { return Row[i]; }
    }

    public int FieldCount
    {
        get { return Row.Table.Columns.Count; }
    }

    public bool GetBoolean(int i)
    {
        return Convert.ToBoolean(Row[i]);
    }

    public byte GetByte(int i)
    {
        return Convert.ToByte(Row[i]);
    }

    public long GetBytes(int i, long fieldOffset, byte[] buffer, int bufferoffset, int length)
    {
        throw new NotSupportedException("GetBytes is not supported.");
    }

    public char GetChar(int i)
    {
        return Convert.ToChar(Row[i]);
    }

    public long GetChars(int i, long fieldoffset, char[] buffer, int bufferoffset, int length)
    {
        throw new NotSupportedException("GetChars is not supported.");
    }

    public IDataReader GetData(int i)
    {
        throw new NotSupportedException("GetData is not supported.");
    }

    public string GetDataTypeName(int i)
    {
        return Row[i].GetType().Name;
    }

    public DateTime GetDateTime(int i)
    {
        return Convert.ToDateTime(Row[i]);
    }

    public decimal GetDecimal(int i)
    {
        return Convert.ToDecimal(Row[i]);
    }

    public double GetDouble(int i)
    {
        return Convert.ToDouble(Row[i]);
    }

    public Type GetFieldType(int i)
    {
        return Row[i].GetType();
    }

    public float GetFloat(int i)
    {
        return Convert.ToSingle(Row[i]);
    }

    public Guid GetGuid(int i)
    {
        return (Guid)Row[i];
    }

    public short GetInt16(int i)
    {
        return Convert.ToInt16(Row[i]);
    }

    public int GetInt32(int i)
    {
        return Convert.ToInt32(Row[i]);
    }

    public long GetInt64(int i)
    {
        return Convert.ToInt64(Row[i]);
    }

    public string GetName(int i)
    {
        return Row.Table.Columns[i].ColumnName;
    }

    public int GetOrdinal(string name)
    {
        return Row.Table.Columns.IndexOf(name);
    }

    public string GetString(int i)
    {
        return Row[i].ToString();
    }

    public object GetValue(int i)
    {
        return Row[i];
    }

    public int GetValues(object[] values)
    {
        values = Row.ItemArray;
        return Row.ItemArray.GetLength(0);
    }

    public bool IsDBNull(int i)
    {
        return Convert.IsDBNull(Row[i]);
    }

}
