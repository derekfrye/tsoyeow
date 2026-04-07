using FluentAssertions;

namespace ExcelXmlWriter.Tests;

public sealed class StaticFunctionsTests
{
    [Fact]
    public void ApplyOverPunch_decodes_signed_last_character()
    {
        var result = StaticFunctions.applyOverPunch("12A", 100);

        result.IsOverpunchable.Should().BeTrue();
        result.val.Should().BeApproximately(1.21d, 0.0001d);
    }

    [Theory]
    [InlineData("12345", ExcelDataType.Number)]
    [InlineData("2009-12-31", ExcelDataType.Date)]
    [InlineData("123456789012345678901234567890123", ExcelDataType.String)]
    [InlineData("123-", ExcelDataType.String)]
    [InlineData("plain text", ExcelDataType.String)]
    public void ResolveDataType_classifies_values_using_excel_rules(string value, ExcelDataType expectedType)
    {
        StaticFunctions.ResolveDataType(value).Should().Be(expectedType);
    }
}
