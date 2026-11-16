using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using SlingMD.Outlook.Services.Formatting;
using Xunit;

namespace SlingMD.Core.Tests.Services.Formatting
{
    /// <summary>
    /// <see cref="FallbackThreadId"/>: the id <c>ThreadService.GetConversationId</c> returns when the
    /// topic, PR_CONVERSATION_INDEX and subject all fail. It used to be a random GUID segment, so a
    /// retried sling of the same email landed in a new folder.
    /// </summary>
    public class FallbackThreadIdTests
    {
        private const string MessageId = "<CAF=abc123@mail.example.com>";
        private const string EntryId = "00000000A1B2C3D4E5F60718293A4B5C6D7E8F900700";

        [Fact]
        public void FromIdentity_SameEmail_GivesSameIdEveryTime()
        {
            string first = FallbackThreadId.FromIdentity(() => MessageId, () => EntryId);
            string second = FallbackThreadId.FromIdentity(() => MessageId, () => EntryId);

            Assert.NotNull(first);
            Assert.Equal(first, second);
        }

        [Fact]
        public void FromIdentity_IdHasThreadIdShape()
        {
            string id = FallbackThreadId.FromIdentity(() => MessageId, () => EntryId);

            Assert.Equal(20, id.Length);
            Assert.Matches("^[0-9A-F]{20}$", id);
        }

        [Fact]
        public void FromIdentity_DifferentEmails_GetDifferentIds()
        {
            string a = FallbackThreadId.FromIdentity(() => "<a@example.com>", () => null);
            string b = FallbackThreadId.FromIdentity(() => "<b@example.com>", () => null);

            Assert.NotEqual(a, b);
        }

        [Fact]
        public void FromIdentity_PrefersInternetMessageId_OverEntryId()
        {
            // The Message-ID survives a move to another store; the EntryID does not.
            string withEntryA = FallbackThreadId.FromIdentity(() => MessageId, () => "ENTRY-A");
            string withEntryB = FallbackThreadId.FromIdentity(() => MessageId, () => "ENTRY-B");

            Assert.Equal(withEntryA, withEntryB);
            Assert.Equal(FallbackThreadId.Hash(MessageId), withEntryA);
        }

        [Fact]
        public void FromIdentity_MessageIdThrows_FallsBackToEntryId()
        {
            List<string> logged = new List<string>();
            Func<string> throwing = () => { throw new COMException("PropertyAccessor gone"); };

            string first = FallbackThreadId.FromIdentity(throwing, () => EntryId, logged.Add);
            string second = FallbackThreadId.FromIdentity(throwing, () => EntryId);

            Assert.Equal(FallbackThreadId.Hash(EntryId), first);
            Assert.Equal(first, second);
            Assert.Single(logged);
            Assert.Contains("Internet message id", logged[0]);
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("   ")]
        public void FromIdentity_BlankMessageId_FallsBackToEntryId(string blank)
        {
            Assert.Equal(FallbackThreadId.Hash(EntryId), FallbackThreadId.FromIdentity(() => blank, () => EntryId));
        }

        [Fact]
        public void FromIdentity_NothingReadable_ReturnsNull()
        {
            Func<string> throwing = () => { throw new InvalidComObjectException("RCW released"); };

            Assert.Null(FallbackThreadId.FromIdentity(throwing, throwing));
            Assert.Null(FallbackThreadId.FromIdentity(() => null, () => " "));
            Assert.Null(FallbackThreadId.FromIdentity(null, null));
        }

        [Fact]
        public void FromIdentity_IgnoresSurroundingWhitespace()
        {
            Assert.Equal(
                FallbackThreadId.FromIdentity(() => MessageId, null),
                FallbackThreadId.FromIdentity(() => "  " + MessageId + "\r\n", null));
        }
    }
}
