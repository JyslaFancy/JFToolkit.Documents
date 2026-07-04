using JFToolkit.Documents;
using Xunit;

namespace JFToolkit.Documents.Tests;

public class AttributeTests
{
    [Fact]
    public void AiExampleAttribute_StoresValue()
    {
        var prop = typeof(TestModel).GetProperty(nameof(TestModel.Name))!;
        var attr = prop.GetCustomAttributes(typeof(AiExampleAttribute), false)
            .Cast<AiExampleAttribute>()
            .FirstOrDefault();

        Assert.NotNull(attr);
        Assert.Equal("Widget", attr!.Example);
    }

    [Fact]
    public void ColumnAttribute_StoresName()
    {
        var prop = typeof(TestModel).GetProperty(nameof(TestModel.Name))!;
        var attr = prop.GetCustomAttributes(typeof(ColumnAttribute), false)
            .Cast<ColumnAttribute>()
            .FirstOrDefault();

        Assert.NotNull(attr);
        Assert.Equal("Product Name", attr!.Name);
    }

    [Fact]
    public void AiExampleAttribute_OnMethod()
    {
        var method = typeof(TestModel).GetMethod(nameof(TestModel.DoSomething))!;
        var attr = method.GetCustomAttributes(typeof(AiExampleAttribute), false)
            .Cast<AiExampleAttribute>()
            .FirstOrDefault();

        Assert.NotNull(attr);
        Assert.Equal("example usage", attr!.Example);
    }

    [Fact]
    public void AiExampleAttribute_OnClass()
    {
        var attr = typeof(TestModel).GetCustomAttributes(typeof(AiExampleAttribute), false)
            .Cast<AiExampleAttribute>()
            .FirstOrDefault();

        Assert.NotNull(attr);
        Assert.Equal("product data", attr!.Example);
    }

    [Fact]
    public void ColumnAttribute_WithoutAttribute_UsesPropertyName()
    {
        var prop = typeof(TestModel).GetProperty(nameof(TestModel.Price))!;
        var attr = prop.GetCustomAttributes(typeof(ColumnAttribute), false)
            .Cast<ColumnAttribute>()
            .FirstOrDefault();

        Assert.Null(attr); // No [Column] on Price
    }

    [AiExample("product data")]
    public class TestModel
    {
        [Column("Product Name")]
        [AiExample("Widget")]
        public string Name { get; set; } = "";

        public decimal Price { get; set; }

        [AiExample("example usage")]
        public void DoSomething() { }
    }
}
