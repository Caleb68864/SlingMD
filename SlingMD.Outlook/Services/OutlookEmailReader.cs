using System.Collections.Generic;
using Microsoft.Office.Interop.Outlook;
using SlingMD.Outlook.Models;
using Logger = SlingMD.Outlook.Helpers.Logger;
using SlingMD.Outlook.Infrastructure;

namespace SlingMD.Outlook.Services
{
    /// <summary>
    /// Copies what the note engine needs out of an Outlook <see cref="MailItem"/> into an
    /// <see cref="EmailData"/>, releasing every COM object it opens. This is the only place the email
    /// path reads the <c>MailItem</c> for note content; <see cref="EmailNoteEngine"/> (SlingMD.Core)
    /// does the rest. The reads were moved here from <c>EmailProcessor</c> unchanged.
    /// </summary>
    public static class OutlookEmailReader
    {
        /// <summary>
        /// Reads the subject, sender, received time, bodies, ids and first To recipient.
        /// <paramref name="recipients"/> is the list from <see cref="ReadRecipients"/>, read earlier.
        /// </summary>
        public static EmailData Read(MailItem mail, ContactService contactService, List<EmailRecipientData> recipients)
        {
            EmailData email = new EmailData
            {
                Subject = mail.Subject,
                SenderName = mail.SenderName,
                SenderEmail = contactService.GetSenderEmail(mail),
                ReceivedTime = mail.ReceivedTime,
                HtmlBody = mail.HTMLBody,
                Body = mail.Body,
                FirstToRecipientName = GetFirstRecipient(mail),
                Recipients = recipients ?? new List<EmailRecipientData>()
            };

            var (internetMessageId, entryId) = ExtractEmailUniqueIds(mail);
            email.InternetMessageId = internetMessageId;
            email.EntryId = entryId;
            return email;
        }

        /// <summary>
        /// Every recipient's display name and SMTP address (via
        /// <see cref="ContactService.GetSMTPEmailAddress"/>), in Outlook's order.
        /// </summary>
        public static List<EmailRecipientData> ReadRecipients(MailItem mail, ContactService contactService)
        {
            List<EmailRecipientData> result = new List<EmailRecipientData>();
            // Properly release COM objects to prevent memory leaks
            Recipients recipients = null;
            try
            {
                recipients = mail.Recipients;
                foreach (Recipient recipient in recipients)
                {
                    try
                    {
                        string name = recipient.Name;
                        result.Add(new EmailRecipientData { Name = name, Email = contactService.GetSMTPEmailAddress(recipient) });
                    }
                    finally
                    {
                        // Release each Recipient COM object
                        if (recipient != null)
                        {
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(recipient);
                        }
                    }
                }
            }
            finally
            {
                // Release the Recipients collection COM object
                if (recipients != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(recipients);
                }
            }

            return result;
        }

        private static string GetFirstRecipient(MailItem mail)
        {
            Recipients recipients = null;
            try
            {
                recipients = mail.Recipients;
                foreach (Recipient recipient in recipients)
                {
                    try
                    {
                        if (recipient.Type == (int)OlMailRecipientType.olTo)
                        {
                            string recipientName = recipient.Name;
                            return recipientName;
                        }
                    }
                    finally
                    {
                        // Release each Recipient COM object
                        if (recipient != null)
                        {
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(recipient);
                        }
                    }
                }
            }
            finally
            {
                // Release the Recipients collection COM object
                if (recipients != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(recipients);
                }
            }
            return "Unknown";
        }

        /// <summary>
        /// Extracts the InternetMessageID and EntryID from a MailItem.
        /// Returns (internetMessageId, entryId).
        /// </summary>
        private static (string internetMessageId, string entryId) ExtractEmailUniqueIds(MailItem mail)
        {
            string entryId = mail.EntryID;
            string internetMessageId = null;
            PropertyAccessor pa = null;
            try
            {
                // Try to get InternetMessageID via PropertyAccessor (works for most accounts).
                // The accessor is a COM object that must be released or it leaks per slung email.
                pa = mail.PropertyAccessor;
                internetMessageId = pa.GetProperty(MapiPropertyTags.PrInternetMessageId) as string;
            }
            catch (System.Exception ex)
            {
                Logger.Instance.Debug($"PropertyAccessor method failed for InternetMessageID: {ex.Message}");
            }
            finally
            {
                if (pa != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(pa);
            }
            // Fallback to property if available
            if (string.IsNullOrEmpty(internetMessageId))
            {
                try
                {
                    internetMessageId = mail.GetType().GetProperty("InternetMessageID")?.GetValue(mail) as string;
                }
                catch (System.Exception ex)
                {
                    Logger.Instance.Debug($"Reflection method failed for InternetMessageID: {ex.Message}");
                }
            }
            return (internetMessageId, entryId);
        }
    }
}
