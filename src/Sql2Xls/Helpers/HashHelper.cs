using System;
using System.Security.Cryptography;
using System.Text;

namespace Sql2Xls.Helpers;

public static class HashHelper
{
    public static byte[] ComputePasswordHash(string password, byte[] saltValue, int spinCount)
    {
        var hasher = SHA512.Create();
        var hash = hasher.ComputeHash(saltValue.Concat(Encoding.Unicode.GetBytes(password)).ToArray());
        for (var i = 0; i < spinCount; i++)
        {
            var iterator = BitConverter.GetBytes(i);
            if (!BitConverter.IsLittleEndian)
                Array.Reverse(iterator);
            hash = hasher.ComputeHash(hash.Concat(iterator).ToArray());
        }
        return hash;
    }

    public static string HexPasswordConversion(string password)
    {
        byte[] passwordcharacters = System.Text.Encoding.ASCII.GetBytes(password);
        int hash = 0;
        if (passwordcharacters.Length > 0)
        {
            int charindex = passwordcharacters.Length;

            while (charindex-- > 0)
            {
                hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7fff);
                hash ^= passwordcharacters[charindex];
            }
            // main difference from spec, also hash with charcount
            hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7fff);
            hash ^= passwordcharacters.Length;
            hash ^= (0x8000 | ('n' << 8) | 'k');
        }

        return Convert.ToString(hash, 16).ToUpperInvariant();
    }
}
