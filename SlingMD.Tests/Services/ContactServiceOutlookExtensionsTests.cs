using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.Office.Interop.Outlook;
using Moq;
using SlingMD.Outlook.Models;
using SlingMD.Outlook.Services;
using SlingMD.Tests.Infrastructure;
using Xunit;

namespace SlingMD.Tests.Services
{
    /// <summary>
    /// <c>ExtractContactData</c> on a Moq <see cref="ContactItem"/>: the add-in reads the contact
    /// (<see cref="ContactServiceOutlookExtensions"/>) and SlingMD.Core builds the context. Every member
    /// the reader uses is set up, because SlingMD.Tests embeds the interop types and a proxy lacks any
    /// member this assembly never references (see AppointmentProcessorExportTests). Passed on the old,
    /// single-class ContactService too (the call compiles against either).
    /// </summary>
    public class ContactServiceOutlookExtensionsTests : IDisposable
    {
        private readonly string _testDir;
        private readonly ContactService _service;

        public ContactServiceOutlookExtensionsTests()
        {
            _testDir = Path.Combine(Path.GetTempPath(), "SlingMDTests", "ContactExtract", Guid.NewGuid().ToString("N"));
            ObsidianSettings settings = new ObsidianSettings { VaultBasePath = _testDir, VaultName = "Vault", ContactsFolder = "Contacts" };
            FileService files = new FileService(settings);
            _service = new ContactService(files, new TemplateService(files), new FakeClock(new DateTime(2026, 9, 23, 9, 0, 0)));
        }

        public void Dispose()
        {
            try
            {
                Directory.Delete(_testDir, true);
            }
            catch (IOException)
            {
            }
        }

        private static Mock<ContactItem> Contact()
        {
            Mock<ContactItem> contact = new Mock<ContactItem>();
            contact.Setup(c => c.FullName).Returns(string.Empty);
            contact.Setup(c => c.FirstName).Returns(string.Empty);
            contact.Setup(c => c.LastName).Returns(string.Empty);
            contact.Setup(c => c.FileAs).Returns(string.Empty);
            contact.Setup(c => c.BusinessTelephoneNumber).Returns(string.Empty);
            contact.Setup(c => c.MobileTelephoneNumber).Returns(string.Empty);
            contact.Setup(c => c.HomeTelephoneNumber).Returns(string.Empty);
            contact.Setup(c => c.Email1Address).Returns(string.Empty);
            contact.Setup(c => c.CompanyName).Returns(string.Empty);
            contact.Setup(c => c.JobTitle).Returns(string.Empty);
            contact.Setup(c => c.BusinessAddress).Returns(string.Empty);
            contact.Setup(c => c.HomeAddress).Returns(string.Empty);
            contact.Setup(c => c.Birthday).Returns(new DateTime(4501, 1, 1));
            contact.Setup(c => c.Body).Returns(string.Empty);
            return contact;
        }

        [Fact]
        public void ExtractContactData_FullContact()
        {
            Mock<ContactItem> contact = Contact();
            contact.Setup(c => c.FullName).Returns("Jane Q. Doe");
            contact.Setup(c => c.BusinessTelephoneNumber).Returns("555-0100");
            contact.Setup(c => c.Email1Address).Returns("jane@example.com");
            contact.Setup(c => c.CompanyName).Returns("Contoso");
            contact.Setup(c => c.JobTitle).Returns("CFO");
            contact.Setup(c => c.BusinessAddress).Returns("1 Main St");
            contact.Setup(c => c.Birthday).Returns(new DateTime(1980, 5, 4));
            contact.Setup(c => c.Body).Returns("Met at the conference");

            ContactTemplateContext context = _service.ExtractContactData(contact.Object);

            Assert.Equal(new[] { "title", "type", "created", "tags", "company", "email" }, context.Metadata.Keys.ToArray());
            Assert.Equal("Jane Q. Doe", context.ContactName);
            Assert.Equal("JaneD", context.ContactShortName);
            Assert.Equal("2026-09-23", context.Created);
            Assert.Equal("555-0100", context.Phone);
            Assert.Equal("1980-05-04", context.Birthday);
            Assert.Equal("Met at the conference", context.Notes);
            Assert.Equal("Jane Q. Doe.md", context.FileName);
        }

        [Fact]
        public void ExtractContactData_Fallbacks_NameFromPartsMobilePhoneHomeAddressNoBirthday()
        {
            Mock<ContactItem> contact = Contact();
            contact.Setup(c => c.FirstName).Returns("Sam");
            contact.Setup(c => c.LastName).Returns("Lee");
            contact.Setup(c => c.MobileTelephoneNumber).Returns("555-0199");
            contact.Setup(c => c.HomeAddress).Returns("2 Side St");

            ContactTemplateContext context = _service.ExtractContactData(contact.Object);

            Assert.Equal("Sam Lee", context.ContactName);
            Assert.Equal("555-0199", context.Phone);
            Assert.Equal("2 Side St", context.Address);
            Assert.Equal(string.Empty, context.Birthday);
            Assert.Equal(new[] { "title", "type", "created", "tags" }, context.Metadata.Keys.ToArray());
        }

        [Fact]
        public void ExtractContactData_NothingReadable_UnknownContact()
        {
            Mock<ContactItem> contact = Contact();
            contact.Setup(c => c.FullName).Throws(new System.Runtime.InteropServices.COMException("gone"));

            Assert.Equal("Unknown Contact", _service.ExtractContactData(contact.Object).ContactName);
        }
    }
}
