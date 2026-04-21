using Microsoft.Office.Interop.Outlook;
using SlingMD.Outlook.Services;
using Xunit;

namespace SlingMD.Tests.Services
{
    public class FlagMonitorServiceTests
    {
        [Fact]
        public void HasFlagTransitioned_NoFlagToMarked_ReturnsTrue()
        {
            bool result = FlagMonitorService.HasFlagTransitioned(
                OlFlagStatus.olNoFlag,
                OlFlagStatus.olFlagMarked);

            Assert.True(result);
        }

        [Fact]
        public void HasFlagTransitioned_MarkedToMarked_ReturnsFalse()
        {
            bool result = FlagMonitorService.HasFlagTransitioned(
                OlFlagStatus.olFlagMarked,
                OlFlagStatus.olFlagMarked);

            Assert.False(result);
        }

        [Fact]
        public void HasFlagTransitioned_MarkedToComplete_ReturnsFalse()
        {
            bool result = FlagMonitorService.HasFlagTransitioned(
                OlFlagStatus.olFlagMarked,
                OlFlagStatus.olFlagComplete);

            Assert.False(result);
        }

        [Fact]
        public void HasFlagTransitioned_NoFlagToNoFlag_ReturnsFalse()
        {
            bool result = FlagMonitorService.HasFlagTransitioned(
                OlFlagStatus.olNoFlag,
                OlFlagStatus.olNoFlag);

            Assert.False(result);
        }

        [Fact]
        public void HasFlagTransitioned_CompleteToMarked_ReturnsTrue()
        {
            bool result = FlagMonitorService.HasFlagTransitioned(
                OlFlagStatus.olFlagComplete,
                OlFlagStatus.olFlagMarked);

            Assert.True(result);
        }

        [Fact]
        public void HasFlagTransitioned_CompleteToNoFlag_ReturnsFalse()
        {
            bool result = FlagMonitorService.HasFlagTransitioned(
                OlFlagStatus.olFlagComplete,
                OlFlagStatus.olNoFlag);

            Assert.False(result);
        }
    }
}
