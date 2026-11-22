using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Outlook;
using SlingMD.Outlook.Helpers;

namespace SlingMD.Outlook.Services
{
    /// <summary>
    /// The COM half of <see cref="ContactService"/> (SlingMD.Core): reading sender and recipient
    /// addresses and names from Outlook items, and an Outlook <see cref="ContactItem"/>. Extension
    /// methods, so callers still write <c>contactService.BuildEmailList(mail.Recipients, ...)</c>.
    /// Moved here from <c>ContactService</c> unchanged.
    /// </summary>
    public static class ContactServiceOutlookExtensions
    {
        /// <summary>
        /// Attempts to resolve the SMTP address for the sender of <paramref name="mail"/>.
        /// Falls back to <see cref="MailItem.SenderEmailAddress"/> when the property accessor fails.
        /// </summary>
        public static string GetSenderEmail(this ContactService service, MailItem mail)
        {
            return SmtpAddressResolver.ForSender(mail);
        }

        /// <summary>
        /// Builds a list of Obsidian wiki-links (e.g. <c>[[Jane Doe]]</c>) for the chosen recipient type.
        /// </summary>
        public static List<string> BuildLinkedNames(this ContactService service, Recipients recipients, OlMailRecipientType type)
        {
            List<string> names = new List<string>();
            foreach (Recipient recipient in recipients)
            {
                try
                {
                    if (recipient.Type == (int)type)
                    {
                        string email = null;
                        try
                        {
                            email = service.GetSMTPEmailAddress(recipient);
                        }
                        catch (System.Exception ex)
                        {
                            Logger.Instance.Warning($"ContactService.BuildLinkedNames: GetSMTPEmailAddress failed for recipient '{recipient.Name}': {ex.Message}");
                        }
                        names.Add(service.FormatContactLink(recipient.Name, email));
                    }
                }
                finally
                {
                    if (recipient != null)
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(recipient);
                    }
                }
            }

            return names;
        }

        /// <summary>
        /// Collects plain email addresses for recipients of the specified <paramref name="type"/>.
        /// </summary>
        public static List<string> BuildEmailList(this ContactService service, Recipients recipients, OlMailRecipientType type)
        {
            List<string> emails = new List<string>();
            foreach (Recipient recipient in recipients)
            {
                try
                {
                    if (recipient.Type == (int)type)
                    {
                        // ForRecipient releases the PropertyAccessor, falls back to Address when the
                        // SMTP property is empty or unreadable, and never throws for a COM failure.
                        string email = SmtpAddressResolver.ForRecipient(recipient);
                        if (!string.IsNullOrEmpty(email))
                        {
                            emails.Add(email);
                        }
                    }
                }
                finally
                {
                    if (recipient != null)
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(recipient);
                    }
                }
            }

            return emails;
        }

        /// <summary>
        /// Attempts to resolve the SMTP address for a meeting <paramref name="recipient"/>.
        /// Falls back to <see cref="Recipient.Address"/> when the property accessor fails.
        /// </summary>
        public static string GetSMTPEmailAddress(this ContactService service, Recipient recipient)
        {
            return SmtpAddressResolver.ForRecipient(recipient);
        }

        /// <summary>
        /// Builds a list of Obsidian wiki-links (e.g. <c>[[Jane Doe]]</c>) filtered by one or more
        /// <see cref="OlMeetingRecipientType"/> values.
        /// </summary>
        public static List<string> BuildLinkedNames(this ContactService service, Recipients recipients, params OlMeetingRecipientType[] types)
        {
            List<string> linkedNames = new List<string>();
            HashSet<int> typeSet = new HashSet<int>();
            foreach (OlMeetingRecipientType type in types)
            {
                typeSet.Add((int)type);
            }

            foreach (Recipient recipient in recipients)
            {
                try
                {
                    if (typeSet.Contains(recipient.Type))
                    {
                        string name = recipient.Name;
                        if (!string.IsNullOrEmpty(name))
                        {
                            string email = null;
                            try
                            {
                                email = service.GetSMTPEmailAddress(recipient);
                            }
                            catch (System.Exception ex)
                            {
                                Logger.Instance.Warning($"ContactService.BuildLinkedNames(meeting): GetSMTPEmailAddress failed for recipient '{name}': {ex.Message}");
                            }
                            linkedNames.Add(service.FormatContactLink(name, email));
                        }
                    }
                }
                finally
                {
                    if (recipient != null)
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(recipient);
                    }
                }
            }

            return linkedNames;
        }

        /// <summary>
        /// Collects plain email addresses for recipients matching the specified meeting role types.
        /// </summary>
        public static List<string> BuildEmailList(this ContactService service, Recipients recipients, IEnumerable<OlMeetingRecipientType> types)
        {
            List<string> emails = new List<string>();
            HashSet<int> typeSet = new HashSet<int>();
            foreach (OlMeetingRecipientType type in types)
            {
                typeSet.Add((int)type);
            }

            foreach (Recipient recipient in recipients)
            {
                try
                {
                    if (typeSet.Contains(recipient.Type))
                    {
                        // Also releases the PropertyAccessor, which this path used to leak.
                        string email = SmtpAddressResolver.ForRecipient(recipient);
                        if (!string.IsNullOrEmpty(email))
                        {
                            emails.Add(email);
                        }
                    }
                }
                finally
                {
                    if (recipient != null)
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(recipient);
                    }
                }
            }

            return emails;
        }

        /// <summary>
        /// Extracts conference room and equipment names from <see cref="OlMeetingRecipientType.olResource"/>
        /// recipients.
        /// </summary>
        public static List<string> GetMeetingResourceData(this ContactService service, Recipients recipients)
        {
            List<string> resources = new List<string>();
            foreach (Recipient recipient in recipients)
            {
                try
                {
                    if (recipient.Type == (int)OlMeetingRecipientType.olResource)
                    {
                        string name = recipient.Name;
                        if (!string.IsNullOrEmpty(name))
                        {
                            resources.Add(name);
                        }
                    }
                }
                finally
                {
                    if (recipient != null)
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(recipient);
                    }
                }
            }

            return resources;
        }

        /// <summary>
        /// Extracts rich contact data from an Outlook <see cref="ContactItem"/> and returns a fully
        /// populated <see cref="ContactTemplateContext"/>. Each COM property read is individually
        /// try/caught to tolerate missing or restricted properties.
        /// </summary>
        public static ContactTemplateContext ExtractContactData(this ContactService service, ContactItem contact)
        {
            string fullName = string.Empty;
            try
            {
                fullName = contact.FullName ?? string.Empty;
            }
            catch (System.Exception ex)
            {
                Logger.Instance.Warning($"ContactService.ExtractContactData: read FullName failed: {ex.Message}");
            }

            if (string.IsNullOrWhiteSpace(fullName))
            {
                try
                {
                    string lastName = contact.LastName ?? string.Empty;
                    string firstName = contact.FirstName ?? string.Empty;
                    if (!string.IsNullOrWhiteSpace(lastName) || !string.IsNullOrWhiteSpace(firstName))
                    {
                        fullName = string.IsNullOrWhiteSpace(firstName)
                            ? lastName
                            : string.IsNullOrWhiteSpace(lastName)
                                ? firstName
                                : $"{firstName} {lastName}";
                    }
                }
                catch (System.Exception ex)
                {
                    Logger.Instance.Warning($"ContactService.ExtractContactData: read FirstName/LastName failed: {ex.Message}");
                }
            }

            if (string.IsNullOrWhiteSpace(fullName))
            {
                try
                {
                    fullName = contact.FileAs ?? string.Empty;
                }
                catch (System.Exception ex)
                {
                    Logger.Instance.Warning($"ContactService.ExtractContactData: read FileAs failed: {ex.Message}");
                }
            }

            if (string.IsNullOrWhiteSpace(fullName))
            {
                fullName = "Unknown Contact";
            }

            string phone = string.Empty;
            try
            {
                phone = contact.BusinessTelephoneNumber ?? string.Empty;
            }
            catch (System.Exception ex)
            {
                Logger.Instance.Warning($"ContactService.ExtractContactData: read BusinessTelephoneNumber failed: {ex.Message}");
            }

            if (string.IsNullOrWhiteSpace(phone))
            {
                try
                {
                    phone = contact.MobileTelephoneNumber ?? string.Empty;
                }
                catch (System.Exception ex)
                {
                    Logger.Instance.Warning($"ContactService.ExtractContactData: read MobileTelephoneNumber failed: {ex.Message}");
                }
            }

            if (string.IsNullOrWhiteSpace(phone))
            {
                try
                {
                    phone = contact.HomeTelephoneNumber ?? string.Empty;
                }
                catch (System.Exception ex)
                {
                    Logger.Instance.Warning($"ContactService.ExtractContactData: read HomeTelephoneNumber failed: {ex.Message}");
                }
            }

            string email = string.Empty;
            try
            {
                email = contact.Email1Address ?? string.Empty;
            }
            catch (System.Exception ex)
            {
                Logger.Instance.Warning($"ContactService.ExtractContactData: read Email1Address failed: {ex.Message}");
            }

            string company = string.Empty;
            try
            {
                company = contact.CompanyName ?? string.Empty;
            }
            catch (System.Exception ex)
            {
                Logger.Instance.Warning($"ContactService.ExtractContactData: read CompanyName failed: {ex.Message}");
            }

            string jobTitle = string.Empty;
            try
            {
                jobTitle = contact.JobTitle ?? string.Empty;
            }
            catch (System.Exception ex)
            {
                Logger.Instance.Warning($"ContactService.ExtractContactData: read JobTitle failed: {ex.Message}");
            }

            string address = string.Empty;
            try
            {
                address = contact.BusinessAddress ?? string.Empty;
            }
            catch (System.Exception ex)
            {
                Logger.Instance.Warning($"ContactService.ExtractContactData: read BusinessAddress failed: {ex.Message}");
            }

            if (string.IsNullOrWhiteSpace(address))
            {
                try
                {
                    address = contact.HomeAddress ?? string.Empty;
                }
                catch (System.Exception ex)
                {
                    Logger.Instance.Warning($"ContactService.ExtractContactData: read HomeAddress failed: {ex.Message}");
                }
            }

            string birthday = string.Empty;
            try
            {
                DateTime birthdayDate = contact.Birthday;
                if (birthdayDate.Year != 4501)
                {
                    birthday = birthdayDate.ToString("yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture);
                }
            }
            catch (System.Exception ex)
            {
                Logger.Instance.Warning($"ContactService.ExtractContactData: read Birthday failed: {ex.Message}");
            }

            string notes = string.Empty;
            try
            {
                notes = contact.Body ?? string.Empty;
            }
            catch (System.Exception ex)
            {
                Logger.Instance.Warning($"ContactService.ExtractContactData: read Body failed: {ex.Message}");
            }

            return service.BuildContactTemplateContext(fullName, phone, email, company, jobTitle, address, birthday, notes);
        }
    }
}
