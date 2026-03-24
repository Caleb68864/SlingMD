using System;
using System.Collections.Generic;
using Xunit;
using SlingMD.Outlook.Helpers;

namespace SlingMD.Tests.Helpers
{
    public class BoundedHashSetTests
    {
        [Fact]
        public void Add_NewItem_ReturnsTrue()
        {
            // Arrange
            BoundedHashSet set = new BoundedHashSet(100);

            // Act
            bool result = set.Add("item1");

            // Assert
            Assert.True(result);
        }

        [Fact]
        public void Add_DuplicateItem_ReturnsFalse()
        {
            // Arrange
            BoundedHashSet set = new BoundedHashSet(100);
            set.Add("item1");

            // Act
            bool result = set.Add("item1");

            // Assert
            Assert.False(result);
        }

        [Fact]
        public void Contains_ExistingItem_ReturnsTrue()
        {
            // Arrange
            BoundedHashSet set = new BoundedHashSet(100);
            set.Add("item1");

            // Act / Assert
            Assert.True(set.Contains("item1"));
        }

        [Fact]
        public void Contains_NonExistingItem_ReturnsFalse()
        {
            // Arrange
            BoundedHashSet set = new BoundedHashSet(100);

            // Act / Assert
            Assert.False(set.Contains("x"));
        }

        [Fact]
        public void Add_ExceedsCapacity_EvictsOldestItem()
        {
            // Arrange
            BoundedHashSet set = new BoundedHashSet(3);
            set.Add("a");
            set.Add("b");
            set.Add("c");

            // Act
            set.Add("d");

            // Assert
            Assert.False(set.Contains("a"));
            Assert.True(set.Contains("d"));
        }

        [Fact]
        public void Add_ExceedsCapacity_PreservesNewestItems()
        {
            // Arrange
            BoundedHashSet set = new BoundedHashSet(3);
            set.Add("a");
            set.Add("b");
            set.Add("c");
            set.Add("d");

            // Assert
            Assert.True(set.Contains("b"));
            Assert.True(set.Contains("c"));
            Assert.True(set.Contains("d"));
        }

        [Fact]
        public void Contains_CaseInsensitiveByDefault()
        {
            // Arrange
            BoundedHashSet set = new BoundedHashSet(100);
            set.Add("TestItem");

            // Act / Assert
            Assert.True(set.Contains("testitem"));
            Assert.True(set.Contains("TESTITEM"));
        }

        [Fact]
        public void Add_CaseInsensitiveDuplicate_ReturnsFalse()
        {
            // Arrange
            BoundedHashSet set = new BoundedHashSet(100);
            set.Add("TestItem");

            // Act
            bool result = set.Add("testitem");

            // Assert
            Assert.False(result);
        }

        [Fact]
        public void Count_ReflectsAddedItems()
        {
            // Arrange
            BoundedHashSet set = new BoundedHashSet(100);

            // Act
            set.Add("item1");
            set.Add("item2");
            set.Add("item3");

            // Assert
            Assert.Equal(3, set.Count);
        }

        [Fact]
        public void Count_DoesNotExceedCapacity()
        {
            // Arrange
            BoundedHashSet set = new BoundedHashSet(2);

            // Act
            set.Add("a");
            set.Add("b");
            set.Add("c");

            // Assert
            Assert.Equal(2, set.Count);
        }

        [Fact]
        public void Clear_RemovesAllItems()
        {
            // Arrange
            BoundedHashSet set = new BoundedHashSet(100);
            set.Add("item1");
            set.Add("item2");
            set.Add("item3");

            // Act
            set.Clear();

            // Assert
            Assert.Equal(0, set.Count);
            Assert.False(set.Contains("item1"));
        }

        [Fact]
        public void Constructor_CustomComparer_IsRespected()
        {
            // Arrange
            BoundedHashSet set = new BoundedHashSet(100, StringComparer.Ordinal);
            set.Add("Test");

            // Act
            bool result = set.Contains("test");

            // Assert
            Assert.False(result);
        }
    }
}
