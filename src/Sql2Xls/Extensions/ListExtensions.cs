﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Sql2Xls.Extensions;

public static class ListExtensions
{
    public static DataTable ToDataTable<T>(this List<T> list)
    {
        var dt = new DataTable();
        foreach (PropertyInfo info in typeof(T).GetProperties())
        {
            dt.Columns.Add(new DataColumn(info.Name, GetNullableType(info.PropertyType)));
        }
        foreach (T t in list)
        {
            DataRow row = dt.NewRow();
            foreach (PropertyInfo info in typeof(T).GetProperties())
            {
                if (!IsNullableType(info.PropertyType))
                    row[info.Name] = info.GetValue(t, null);
                else
                    row[info.Name] = (info.GetValue(t, null) ?? DBNull.Value);
            }
            dt.Rows.Add(row);
        }
        return dt;
    }

    private static Type GetNullableType(Type t)
    {
        Type returnType = t;
        if (t.IsGenericType && t.GetGenericTypeDefinition().Equals(typeof(Nullable<>)))
        {
            returnType = Nullable.GetUnderlyingType(t);
        }
        return returnType;
    }

    private static bool IsNullableType(Type type)
    {
        return (type == typeof(string) ||
                type.IsArray ||
                (type.IsGenericType &&
                 type.GetGenericTypeDefinition().Equals(typeof(Nullable<>))));
    }
}
