using System.Globalization;

namespace LubanExcelMerge.Luban;

public static class LubanKeyValueNormalizer
{
    public static string? Normalize(LubanField field, string? value)
    {
        var typeName = field.TypeName.Split('#')[0].TrimEnd('!');
        // Nullable keys must retain the distinction between null and zero.
        if (typeName.EndsWith('?'))
            return string.IsNullOrEmpty(value) ? value : Normalize(field with { TypeName = typeName[..^1] }, value);
        if (IsBoolean(typeName))
        {
            if (string.IsNullOrEmpty(value))
                return "0";
            if (value == "1" || bool.TryParse(value, out var boolean) && boolean)
                return "1";
            if (value == "0" || bool.TryParse(value, out boolean) && !boolean)
                return "0";
            return value;
        }

        if (!IsNumeric(typeName))
            return value;
        if (string.IsNullOrEmpty(value))
            return "0";
        return decimal.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var number)
            ? number.ToString("G29", CultureInfo.InvariantCulture)
            : value;
    }

    private static bool IsBoolean(string typeName) =>
        string.Equals(typeName, "bool", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(typeName, "boolean", StringComparison.OrdinalIgnoreCase);

    private static bool IsNumeric(string typeName)
    {
        var simpleName = typeName.Split('.').LastOrDefault() ?? typeName;
        return simpleName.Equals("FPString", StringComparison.OrdinalIgnoreCase) ||
               simpleName is "byte" or "sbyte" or "short" or "ushort" or "int" or "uint" or
                   "long" or "ulong" or "float" or "double" or "decimal";
    }
}
