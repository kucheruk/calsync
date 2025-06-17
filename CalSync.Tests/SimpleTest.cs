using Xunit;

namespace CalSync.Tests;

public class SimpleTest
{
    [Fact]
    public void SimpleTest_ShouldPass()
    {
        // Arrange
        var expected = true;

        // Act
        var actual = true;

        // Assert
        Assert.Equal(expected, actual);
    }
}