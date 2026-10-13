using System.Xml;
using System.Xml.Linq;

namespace LubanExcelMerge.Luban;

/// <summary>Resolves external bean types without guessing Excel column positions.</summary>
public sealed class XmlBeanCatalog
{
    private readonly Dictionary<string, XElement> _beans = new(StringComparer.Ordinal);

    public static XmlBeanCatalog Load(string directory)
    {
        if (!Directory.Exists(directory))
            throw new FormatException($"XML 字段定义目录不存在：{directory}。");
        var catalog = new XmlBeanCatalog();
        foreach (var path in Directory.EnumerateFiles(directory, "*.xml", SearchOption.AllDirectories).Order(StringComparer.Ordinal))
        {
            using var reader = XmlReader.Create(path, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null });
            var document = XDocument.Load(reader);
            foreach (var bean in document.Descendants("bean"))
            {
                var name = string.Join(".", bean.Ancestors("module").Reverse()
                    .Select(module => (string?)module.Attribute("name"))
                    .Append((string?)bean.Attribute("name")).Where(part => !string.IsNullOrEmpty(part)));
                if (!catalog._beans.TryAdd(name, bean))
                    throw new FormatException($"XML bean {name} 重复定义（{path}）。");
            }
        }
        return catalog;
    }

    public IReadOnlyList<LubanField> Resolve(LogicalTableDefinition table)
    {
        var moduleEnd = table.FullName.LastIndexOf('.');
        var qualifiedName = table.ValueType.Contains('.') || moduleEnd < 0
            ? table.ValueType : table.FullName[..(moduleEnd + 1)] + table.ValueType;
        if (!_beans.TryGetValue(qualifiedName, out var bean))
            throw new FormatException($"逻辑表 {table.FullName} 对应的 XML bean {qualifiedName} 未找到。请检查 Defines 和 __tables__.csv。");
        if (bean.Attribute("parent") is not null)
            throw new FormatException($"XML bean {qualifiedName} 使用 parent 属性继承，当前仅支持嵌套 bean 继承。");
        var declarations = bean.Ancestors("bean").Reverse().SelectMany(parent => parent.Elements("var"))
            .Concat(bean.Descendants("var"));
        var fields = declarations.Select((field, index) => new LubanField(index,
            (string?)field.Attribute("name") ?? throw new FormatException($"XML bean {qualifiedName} 存在无名称字段。"),
            (string?)field.Attribute("type") ?? throw new FormatException($"XML bean {qualifiedName} 存在无类型字段。"))).ToArray();
        if (fields.Length == 0 || fields.GroupBy(field => field.Name).Any(group => group.Select(field => field.TypeName).Distinct().Count() > 1))
            throw new FormatException($"XML bean {qualifiedName} 没有字段或继承分支存在同名异类型字段，无法唯一映射。");
        return fields.DistinctBy(field => field.Name).ToArray();
    }

    public LubanSchema Apply(LubanSchema schema, LogicalTableDefinition table)
    {
        var definitions = Resolve(table).ToDictionary(field => field.Name, StringComparer.Ordinal);
        var fields = schema.Fields.Select(field => field.Name is "$type" or "__type__" || field.Name.StartsWith('#')
            ? field with { TypeName = "string" }
            : definitions.TryGetValue(field.Name, out var definition)
            ? field with { TypeName = definition.TypeName }
            : throw new FormatException($"Excel 字段 {field.Name} 不在 XML bean {table.ValueType} 中，无法安全映射。")) .ToArray();
        return schema with { Fields = fields };
    }
}
