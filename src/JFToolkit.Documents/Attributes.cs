namespace JFToolkit.Documents;

/// <summary>
/// Marks a property with an example value to help AI coding agents
/// understand the expected data shape.
/// </summary>
[AttributeUsage(AttributeTargets.Property | AttributeTargets.Method | AttributeTargets.Class, AllowMultiple = false)]
public class AiExampleAttribute : Attribute
{
    /// <summary>An example value for this property/method/class.</summary>
    public string Example { get; }

    public AiExampleAttribute(string example)
    {
        Example = example;
    }
}

/// <summary>
/// Maps a property to a specific column name in tabular formats (CSV, Excel).
/// When not specified, the property name is used.
/// </summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
public class ColumnAttribute : Attribute
{
    /// <summary>The column header name.</summary>
    public string Name { get; }

    public ColumnAttribute(string name)
    {
        Name = name;
    }
}
