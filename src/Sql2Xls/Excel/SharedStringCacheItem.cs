namespace Sql2Xls.Excel;

public readonly struct SharedStringCacheItem
{
    public int Position { get; init; }
    public string Value { get; init; }

    public override string ToString()
    {
        return Value;
    }
}
