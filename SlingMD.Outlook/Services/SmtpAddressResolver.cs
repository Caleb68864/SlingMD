using System;
using Microsoft.Office.Interop.Outlook;
using SlingMD.Outlook.Helpers;
using SlingMD.Outlook.Infrastructure;

namespace SlingMD.Outlook.Services
{
    /// <summary>
    /// Reads SMTP addresses from Outlook items. Exchange senders and recipients otherwise come back as
    /// X.500 paths ("/O=EXCHANGELABS/OU=…"); the PR_SMTP_ADDRESS property holds the real address.
    /// Every PropertyAccessor is released, and failures fall back to the address Outlook shows.
    /// </summary>
    public static class SmtpAddressResolver
    {
        public static string ForSender(MailItem mail)
        {
            // mail.PropertyAccessor mints a COM object that must be released, or one leaks per email.
            // The fallback read is also guarded so a COMException can't escape and abort an export.
            PropertyAccessor pa = null;
            try
            {
                pa = mail.PropertyAccessor;
                return (string)pa.GetProperty(MapiPropertyTags.PrSmtpAddress);
            }
            catch (System.Exception ex)
            {
                Logger.Instance.Warning($"SmtpAddressResolver.ForSender: PropertyAccessor failed, falling back to SenderEmailAddress: {ex.Message}");
                return SafeComAction.Execute(
                    () => mail.SenderEmailAddress ?? string.Empty,
                    "SmtpAddressResolver.ForSender: SenderEmailAddress fallback",
                    string.Empty);
            }
            finally
            {
                if (pa != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(pa);
            }
        }

        public static string ForRecipient(Recipient recipient)
        {
            PropertyAccessor pa = null;
            try
            {
                pa = recipient.PropertyAccessor;
                return pa.GetProperty(MapiPropertyTags.PrSmtpAddress) as string ?? recipient.Address;
            }
            catch
            {
                return recipient.Address;
            }
            finally
            {
                // Release the PropertyAccessor COM object (distinct from recipient) to avoid a
                // per-recipient leak.
                if (pa != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(pa);
            }
        }
    }
}
