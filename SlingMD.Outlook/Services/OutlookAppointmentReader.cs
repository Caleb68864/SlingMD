using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Outlook;
using SlingMD.Outlook.Models;
using Logger = SlingMD.Outlook.Helpers.Logger;

namespace SlingMD.Outlook.Services
{
    /// <summary>
    /// Copies what the note engine needs out of an Outlook <see cref="AppointmentItem"/> into an
    /// <see cref="AppointmentData"/>, releasing every COM object it opens. Each read is guarded on its
    /// own, so one unreadable property leaves its default and the rest are still read.
    /// <see cref="AppointmentNoteEngine"/> (SlingMD.Core) does the rest. The reads were moved here
    /// from <c>AppointmentProcessor</c> unchanged.
    /// </summary>
    public static class OutlookAppointmentReader
    {
        /// <summary>
        /// Reads subject, body, location, start/end (default <paramref name="now"/>), recurrence state,
        /// GlobalAppointmentID, calendar name, organizer and the recipients split by type.
        /// </summary>
        public static AppointmentData Read(AppointmentItem appointment, DateTime now)
        {
            string subject = string.Empty;
            string body = string.Empty;
            string location = string.Empty;
            DateTime startTime = now;
            DateTime endTime = now;
            string organizerName = string.Empty;
            string organizerEmail = string.Empty;
            string globalAppointmentId = string.Empty;
            OlRecurrenceState recurrenceState = OlRecurrenceState.olApptNotRecurring;
            List<string> requiredAttendees = new List<string>();
            List<string> optionalAttendees = new List<string>();
            List<string> resourceAttendees = new List<string>();

            try { subject = appointment.Subject ?? string.Empty; }
            catch (System.Exception ex) { Logger.Instance.Debug($"Could not read Subject: {ex.Message}"); }

            try { body = appointment.Body ?? string.Empty; }
            catch (System.Exception ex) { Logger.Instance.Debug($"Could not read Body: {ex.Message}"); }

            try { location = appointment.Location ?? string.Empty; }
            catch (System.Exception ex) { Logger.Instance.Debug($"Could not read Location: {ex.Message}"); }

            try { startTime = appointment.Start; }
            catch (System.Exception ex) { Logger.Instance.Debug($"Could not read Start: {ex.Message}"); }

            try { endTime = appointment.End; }
            catch (System.Exception ex) { Logger.Instance.Debug($"Could not read End: {ex.Message}"); }

            try { recurrenceState = appointment.RecurrenceState; }
            catch (System.Exception ex) { Logger.Instance.Debug($"Could not read RecurrenceState: {ex.Message}"); }

            try { globalAppointmentId = appointment.GlobalAppointmentID ?? string.Empty; }
            catch (System.Exception ex) { Logger.Instance.Debug($"Could not read GlobalAppointmentID: {ex.Message}"); }

            string calendarName = GetCalendarName(appointment);

            // Extract organizer via AddressEntry COM object
            AddressEntry organizerEntry = null;
            try
            {
                organizerEntry = appointment.GetOrganizer();
                if (organizerEntry != null)
                {
                    try { organizerName = organizerEntry.Name ?? string.Empty; }
                    catch (System.Exception ex) { Logger.Instance.Debug($"Could not read organizer Name: {ex.Message}"); }

                    try { organizerEmail = organizerEntry.Address ?? string.Empty; }
                    catch (System.Exception ex) { Logger.Instance.Debug($"Could not read organizer Address: {ex.Message}"); }
                }
            }
            catch (System.Exception ex)
            {
                Logger.Instance.Debug($"Could not get organizer: {ex.Message}");
            }
            finally
            {
                if (organizerEntry != null)
                {
                    Marshal.ReleaseComObject(organizerEntry);
                }
            }

            // Extract recipients by type
            Recipients recipients = null;
            try
            {
                recipients = appointment.Recipients;
                foreach (Recipient recipient in recipients)
                {
                    try
                    {
                        string recipientName = string.Empty;
                        OlMeetingRecipientType recipientType = OlMeetingRecipientType.olRequired;

                        try { recipientName = recipient.Name ?? string.Empty; }
                        catch (System.Exception ex) { Logger.Instance.Debug($"Could not read recipient Name: {ex.Message}"); }

                        try { recipientType = (OlMeetingRecipientType)recipient.Type; }
                        catch (System.Exception ex) { Logger.Instance.Debug($"Could not read recipient Type: {ex.Message}"); }

                        if (!string.IsNullOrWhiteSpace(recipientName))
                        {
                            if (recipientType == OlMeetingRecipientType.olOptional)
                            {
                                optionalAttendees.Add(recipientName);
                            }
                            else if (recipientType == OlMeetingRecipientType.olResource)
                            {
                                resourceAttendees.Add(recipientName);
                            }
                            else
                            {
                                requiredAttendees.Add(recipientName);
                            }
                        }
                    }
                    catch (System.Exception ex)
                    {
                        Logger.Instance.Debug($"Could not process recipient: {ex.Message}");
                    }
                    finally
                    {
                        if (recipient != null)
                        {
                            Marshal.ReleaseComObject(recipient);
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                Logger.Instance.Debug($"Could not read Recipients: {ex.Message}");
            }
            finally
            {
                if (recipients != null)
                {
                    Marshal.ReleaseComObject(recipients);
                }
            }

            return new AppointmentData
            {
                Subject = subject,
                Body = body,
                Location = location,
                Start = startTime,
                End = endTime,
                OrganizerName = organizerName,
                OrganizerEmail = organizerEmail,
                GlobalAppointmentId = globalAppointmentId,
                Recurrence = ToRecurrence(recurrenceState),
                CalendarName = calendarName,
                RequiredAttendees = requiredAttendees,
                OptionalAttendees = optionalAttendees,
                ResourceAttendees = resourceAttendees
            };
        }

        /// <summary>Outlook's recurrence state as the Core enum.</summary>
        public static AppointmentRecurrence ToRecurrence(OlRecurrenceState recurrenceState)
        {
            switch (recurrenceState)
            {
                case OlRecurrenceState.olApptMaster:
                    return AppointmentRecurrence.Master;
                case OlRecurrenceState.olApptOccurrence:
                    return AppointmentRecurrence.Occurrence;
                case OlRecurrenceState.olApptException:
                    return AppointmentRecurrence.Exception;
                default:
                    return AppointmentRecurrence.NotRecurring;
            }
        }

        /// <summary>
        /// Name of the calendar holding the appointment, qualified by its store (for example
        /// "Calendar (jane@firm.com)") so the default calendars of different accounts stay distinct.
        /// Empty when Outlook does not expose the folder.
        /// </summary>
        private static string GetCalendarName(AppointmentItem appointment)
        {
            object parent = null;
            object seriesParent = null;
            Store store = null;
            try
            {
                parent = appointment.Parent;

                // An occurrence of a recurring series can report its series master as the parent;
                // the calendar folder is then one level further up.
                AppointmentItem master = parent as AppointmentItem;
                if (master != null)
                {
                    seriesParent = master.Parent;
                }

                MAPIFolder folder = (seriesParent ?? parent) as MAPIFolder;
                if (folder == null)
                {
                    return string.Empty;
                }

                string storeName = string.Empty;
                try
                {
                    store = folder.Store;
                    storeName = store.DisplayName ?? string.Empty;
                }
                catch (System.Exception ex)
                {
                    Logger.Instance.Debug($"Could not read calendar store name: {ex.Message}");
                }

                return CalendarFolderService.FormatCalendarLabel(folder.Name, storeName);
            }
            catch (System.Exception ex)
            {
                Logger.Instance.Debug($"Could not read calendar name: {ex.Message}");
                return string.Empty;
            }
            finally
            {
                if (store != null && Marshal.IsComObject(store)) Marshal.ReleaseComObject(store);
                if (seriesParent != null && Marshal.IsComObject(seriesParent)) Marshal.ReleaseComObject(seriesParent);
                if (parent != null && Marshal.IsComObject(parent)) Marshal.ReleaseComObject(parent);
            }
        }
    }
}
