using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Outlook;
using Moq;
using SlingMD.Outlook.Infrastructure;
using SlingMD.Outlook.Models;
using SlingMD.Outlook.Services;
using SlingMD.Outlook.Services.Formatting;
using Xunit;

namespace SlingMD.Tests.Services
{
    /// <summary>
    /// Tests for <see cref="ConversationIdReader.GetConversationId"/> against Moq doubles of the Outlook
    /// <see cref="MailItem"/> and <see cref="PropertyAccessor"/> interfaces. Moq can proxy these
    /// even though SlingMD.Tests embeds the interop types (checked under .NET 10 and Mono; the
    /// VSTest run on Windows is what confirms it for .NET Framework).
    /// </summary>
    public class ThreadServiceConversationIdTests
    {
        private const string MessageId = "<BN6PR11MB1234@BN6PR11MB1234.namprd11.prod.outlook.com>";
        private const string EntryId = "00000000ABCDEF0123456789";

        private readonly ObsidianSettings _settings;
        private readonly ConversationIdReader _threadService;
        private readonly ThreadIdHasher _hasher;

        public ThreadServiceConversationIdTests()
        {
            _settings = new ObsidianSettings();
            _threadService = new ConversationIdReader(_settings);
            _hasher = new ThreadIdHasher(new SubjectCleanerService(_settings));
        }

        private static COMException ComError()
        {
            return new COMException("The operation failed.", unchecked((int)0x80004005));
        }

        private static Mock<PropertyAccessor> Accessor(object conversationIndex, string internetMessageId)
        {
            Mock<PropertyAccessor> accessor = new Mock<PropertyAccessor>();
            accessor.Setup(a => a.GetProperty(MapiPropertyTags.PrConversationIndex)).Returns(conversationIndex);
            accessor.Setup(a => a.GetProperty(MapiPropertyTags.PrInternetMessageId)).Returns(internetMessageId);
            return accessor;
        }

        [Fact]
        public void GetConversationId_WithTopic_HashesTheTopic()
        {
            Mock<MailItem> mail = new Mock<MailItem>();
            mail.Setup(m => m.ConversationTopic).Returns("Quarterly loan review");
            mail.Setup(m => m.Subject).Returns("RE: Something else entirely");

            string id = _threadService.GetConversationId(mail.Object);

            Assert.Equal(_hasher.Hash("Quarterly loan review"), id);
            Assert.Equal(20, id.Length);
            // The topic wins: the property accessor is never consulted.
            mail.Verify(m => m.PropertyAccessor, Times.Never());
        }

        [Fact]
        public void GetConversationId_SameTopicWithReplyPrefix_GivesTheSameId()
        {
            Mock<MailItem> original = new Mock<MailItem>();
            original.Setup(m => m.ConversationTopic).Returns("Quarterly loan review");
            Mock<MailItem> reply = new Mock<MailItem>();
            reply.Setup(m => m.ConversationTopic).Returns("RE: Quarterly loan review");

            Assert.Equal(
                _threadService.GetConversationId(original.Object),
                _threadService.GetConversationId(reply.Object));
        }

        [Fact]
        public void GetConversationId_NoTopic_UsesTheFirst22BytesOfTheConversationIndex()
        {
            byte[] index = new byte[27];
            for (int i = 0; i < index.Length; i++)
            {
                index[i] = (byte)(i + 1);
            }

            Mock<MailItem> mail = new Mock<MailItem>();
            mail.Setup(m => m.ConversationTopic).Returns(string.Empty);
            mail.Setup(m => m.Subject).Returns("Quarterly loan review");
            mail.Setup(m => m.PropertyAccessor).Returns(Accessor(index, MessageId).Object);

            string id = _threadService.GetConversationId(mail.Object);

            // Hex of bytes 01..0A: the first 20 hex digits of the 22-byte header.
            Assert.Equal("0102030405060708090A", id);
        }

        [Fact]
        public void GetConversationId_NoTopic_ReadsTheConversationIndexAsBinary()
        {
            // PR_CONVERSATION_INDEX is PT_BINARY (0x0102). Outlook returns its byte[] only for the binary
            // tag; the Unicode-string tag (0x001F) used before 2026-11-16 never yielded one, so the
            // conversation-index strategy never ran and these emails were grouped by subject.
            byte[] index = new byte[27];
            for (int i = 0; i < index.Length; i++)
            {
                index[i] = (byte)(0xA0 + i);
            }

            Mock<PropertyAccessor> accessor = new Mock<PropertyAccessor>();
            accessor.Setup(a => a.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x00710102")).Returns(index);
            accessor.Setup(a => a.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x0071001F")).Returns((object)null);
            Mock<MailItem> mail = new Mock<MailItem>();
            mail.Setup(m => m.ConversationTopic).Returns(string.Empty);
            mail.Setup(m => m.Subject).Returns("Quarterly loan review");
            mail.Setup(m => m.PropertyAccessor).Returns(accessor.Object);

            string id = _threadService.GetConversationId(mail.Object);

            Assert.Equal("A0A1A2A3A4A5A6A7A8A9", id);
            Assert.NotEqual(_hasher.Hash("Quarterly loan review"), id);
        }

        [Fact]
        public void GetConversationId_NoTopic_ConversationIndexTooShort_HashesTheSubject()
        {
            Mock<MailItem> mail = new Mock<MailItem>();
            mail.Setup(m => m.ConversationTopic).Returns((string)null);
            mail.Setup(m => m.Subject).Returns("Quarterly loan review");
            mail.Setup(m => m.PropertyAccessor).Returns(Accessor(new byte[10], MessageId).Object);

            Assert.Equal(_hasher.Hash("Quarterly loan review"), _threadService.GetConversationId(mail.Object));
        }

        [Fact]
        public void GetConversationId_NoTopic_ConversationIndexThrows_HashesTheSubject()
        {
            Mock<PropertyAccessor> accessor = new Mock<PropertyAccessor>();
            accessor.Setup(a => a.GetProperty(It.IsAny<string>())).Throws(ComError());
            Mock<MailItem> mail = new Mock<MailItem>();
            mail.Setup(m => m.ConversationTopic).Returns(string.Empty);
            mail.Setup(m => m.Subject).Returns("Quarterly loan review");
            mail.Setup(m => m.PropertyAccessor).Returns(accessor.Object);

            // Falls through to the subject, not to the identity or random fallback.
            Assert.Equal(_hasher.Hash("Quarterly loan review"), _threadService.GetConversationId(mail.Object));
        }

        [Fact]
        public void GetConversationId_TopicUnreadable_HashesTheInternetMessageId()
        {
            Mock<MailItem> mail = new Mock<MailItem>();
            mail.Setup(m => m.ConversationTopic).Throws(ComError());
            mail.Setup(m => m.PropertyAccessor).Returns(Accessor(null, MessageId).Object);
            mail.Setup(m => m.EntryID).Returns(EntryId);

            Assert.Equal(FallbackThreadId.Hash(MessageId), _threadService.GetConversationId(mail.Object));
        }

        [Fact]
        public void GetConversationId_TopicAndMessageIdUnreadable_HashesTheEntryId()
        {
            Mock<PropertyAccessor> accessor = new Mock<PropertyAccessor>();
            accessor.Setup(a => a.GetProperty(It.IsAny<string>())).Throws(ComError());
            Mock<MailItem> mail = new Mock<MailItem>();
            mail.Setup(m => m.ConversationTopic).Throws(ComError());
            mail.Setup(m => m.PropertyAccessor).Returns(accessor.Object);
            mail.Setup(m => m.EntryID).Returns(EntryId);

            Assert.Equal(FallbackThreadId.Hash(EntryId), _threadService.GetConversationId(mail.Object));
        }

        [Fact]
        public void GetConversationId_TopicUnreadable_IsTheSameOnEveryCall()
        {
            Mock<MailItem> mail = new Mock<MailItem>();
            mail.Setup(m => m.ConversationTopic).Throws(new InvalidComObjectException("COM object separated from its RCW"));
            mail.Setup(m => m.PropertyAccessor).Returns(Accessor(null, MessageId).Object);

            // A retried sling of the same email must land in the same thread folder.
            Assert.Equal(_threadService.GetConversationId(mail.Object), _threadService.GetConversationId(mail.Object));
        }

        [Fact]
        public void GetConversationId_NothingReadable_ReturnsARandom20CharacterHexId()
        {
            Mock<MailItem> mail = new Mock<MailItem>();
            mail.Setup(m => m.ConversationTopic).Throws(ComError());
            mail.Setup(m => m.PropertyAccessor).Throws(ComError());
            mail.Setup(m => m.EntryID).Throws(ComError());

            string first = _threadService.GetConversationId(mail.Object);
            string second = _threadService.GetConversationId(mail.Object);

            Assert.Matches("^[0-9A-F]{20}$", first);
            Assert.Matches("^[0-9A-F]{20}$", second);
            Assert.NotEqual(first, second);
        }
    }
}
