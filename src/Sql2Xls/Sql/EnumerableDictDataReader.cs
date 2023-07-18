namespace Sql2Xls.Sql;

public class EnumerableDictDataReader : DictionaryDataReader
{
    private readonly IEnumerator<IDictionary<string, object>> _enumerator;
    private IDictionary<string, object> _current;
    private bool isInitialized = false;

    /// <summary>
    /// Create an IDataReader over an instance of IEnumerable&lt;>.
    /// 
    /// Note: anonymous type arguments are acceptable.
    /// 
    /// Use other constructor for IEnumerable.
    /// </summary>
    /// <param name="collection">IEnumerable&lt;>. For IEnumerable use other constructor and specify type.</param>
    public EnumerableDictDataReader(IEnumerable<IDictionary<string, object>> collection)
    {
        _enumerator = collection.GetEnumerator();
        isInitialized = false;
    }

    /// <summary>
    /// Create an IDataReader over an instance of IEnumerable.
    /// Use other constructor for IEnumerable&lt;>
    /// </summary>
    /// <param name="collection"></param>
    /// <param name="elementType"></param>
    public EnumerableDictDataReader(IEnumerable<IDictionary<string, object>> collection, ICollection<string> fieldNames, ICollection<Type> types)
        : base(fieldNames, types)
    {
        _enumerator = collection.GetEnumerator();
        isInitialized = true;
    }

    /// <summary>
    /// Return the value of the specified field.
    /// </summary>
    /// <returns>
    /// The <see cref="T:System.Object"/> which will contain the field value upon return.
    /// </returns>
    /// <param name="i">The index of the field to find. 
    /// </param><exception cref="T:System.IndexOutOfRangeException">The index passed was outside the range of 0 through <see cref="P:System.Data.IDataRecord.FieldCount"/>. 
    /// </exception><filterpriority>2</filterpriority>
    public override object GetValue(int i)
    {
        if (i < 0 || i >= Fields.Count)
        {
            throw new IndexOutOfRangeException();
        }

        return _current[Fields[i]];
    }

    /// <summary>
    /// Advances the <see cref="T:System.Data.IDataReader"/> to the next record.
    /// </summary>
    /// <returns>
    /// true if there are more rows; otherwise, false.
    /// </returns>
    /// <filterpriority>2</filterpriority>
    public override bool Read()
    {
        bool returnValue = _enumerator.MoveNext();
        if (returnValue)
        {
            _current = _enumerator.Current;
            if (!isInitialized)
            {
                SetFields(_current);
                isInitialized = true;
            }
        }
        else
        {
            _current = null;
        }
        return returnValue;
    }
}
