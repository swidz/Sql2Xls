namespace Sql2Xls.Excel
{
    public sealed class SharedStringCacheItem
    {
        public int Position { get; set; }
        public string Value { get; set; }

        public SharedStringCacheItem(int position, string value)
        {
            Position = position;
            Value = value;
        }

        public static SharedStringCacheItem Create(int position, string value)
        {
            return new SharedStringCacheItem(position, value);
        }

        public override string ToString()
        {
            return Value;
        }
    }
}
