using System;
using System.Runtime.InteropServices;
using Xunit;
using SlingMD.Outlook.Helpers;

namespace SlingMD.Tests.Helpers
{
    public class SafeComActionTests
    {
        [Fact]
        public void Execute_DelegateSucceeds_ReturnsValue()
        {
            // Arrange
            Func<string> action = () => "hello";

            // Act
            string result = SafeComAction.Execute(action, "test", "default");

            // Assert
            Assert.Equal("hello", result);
        }

        [Fact]
        public void Execute_DelegateThrowsCOMException_ReturnsDefault()
        {
            // Arrange
            Func<string> action = () => throw new COMException("test error");

            // Act
            string result = SafeComAction.Execute(action, "test", "fallback");

            // Assert
            Assert.Equal("fallback", result);
        }

        [Fact]
        public void Execute_DelegateThrowsGenericException_ReturnsDefault()
        {
            // Arrange
            Func<string> action = () => throw new InvalidOperationException("test");

            // Act
            string result = SafeComAction.Execute(action, "test", "fallback");

            // Assert
            Assert.Equal("fallback", result);
        }

        [Fact]
        public void Execute_VoidDelegateSucceeds_DoesNotThrow()
        {
            // Arrange
            bool wasExecuted = false;
            Action action = () => { wasExecuted = true; };

            // Act
            SafeComAction.Execute(action, "test");

            // Assert
            Assert.True(wasExecuted);
        }

        [Fact]
        public void Execute_VoidDelegateThrows_DoesNotThrow()
        {
            // Arrange
            Action action = () => throw new COMException("com error");

            // Act
            System.Exception caughtException = null;
            try
            {
                SafeComAction.Execute(action, "test");
            }
            catch (System.Exception ex)
            {
                caughtException = ex;
            }

            // Assert
            Assert.Null(caughtException);
        }

        [Fact]
        public void Execute_ReturnsDefaultForValueType()
        {
            // Arrange
            Func<int> action = () => throw new COMException("com error");

            // Act
            int result = SafeComAction.Execute(action, "test", 0);

            // Assert
            Assert.Equal(0, result);
        }

        [Fact]
        public void ExecuteAndRelease_DelegateSucceeds_ReturnsValue()
        {
            // Arrange
            object expected = new object();
            Func<object> action = () => expected;

            // Act
            object result = SafeComAction.ExecuteAndRelease(action, "test", null);

            // Assert
            Assert.NotNull(result);
            Assert.Same(expected, result);
        }
    }
}
