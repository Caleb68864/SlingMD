using System.Collections.Generic;
using System.Linq;
using System.Text;
using SlingMD.Outlook.Models;

namespace SlingMD.Outlook.Services
{
    /// <summary>
    /// Tallies a batch sling (Sling Multiple, Complete Thread) and writes the summary shown at the
    /// end: how many were saved, how many were already in the vault, and which ones failed and why.
    /// Pure helper with no Outlook or UI dependencies.
    /// </summary>
    public class BatchSlingReport
    {
        /// <summary>How many failures are listed by name before the rest are summarised.</summary>
        public const int MaxListedFailures = 10;

        private readonly List<KeyValuePair<string, string>> _failures = new List<KeyValuePair<string, string>>();

        public BatchSlingReport(int total)
        {
            Total = total;
        }

        public int Total { get; }
        public int Saved { get; private set; }
        public int Duplicates { get; private set; }
        public int Failed => _failures.Count;
        public int Processed => Saved + Duplicates + Failed;

        /// <summary>Whether the user stopped the batch before every email was processed.</summary>
        public bool Cancelled { get; set; }

        /// <summary>Full path of the last note saved, if any.</summary>
        public string LastSavedNotePath { get; private set; }

        public void Add(string subject, EmailSlingResult result)
        {
            switch (result?.Outcome)
            {
                case EmailSlingOutcome.Saved:
                    Saved++;
                    LastSavedNotePath = result.NotePath;
                    break;
                case EmailSlingOutcome.Duplicate:
                    Duplicates++;
                    break;
                default:
                    string error = string.IsNullOrWhiteSpace(result?.ErrorMessage) ? "unknown error" : result.ErrorMessage.Trim();
                    _failures.Add(new KeyValuePair<string, string>(string.IsNullOrWhiteSpace(subject) ? "(no subject)" : subject.Trim(), error));
                    break;
            }
        }

        /// <summary>
        /// Builds the summary, e.g. "Slung 8 of 10 emails. 1 was already in your vault." followed by the
        /// failures, one per line.
        /// </summary>
        /// <param name="itemsLabel">What the batch held, e.g. "emails" or "emails from this thread".</param>
        public string BuildMessage(string itemsLabel)
        {
            StringBuilder message = new StringBuilder();
            message.Append($"Slung {Saved} of {Total} {itemsLabel}.");
            if (Duplicates > 0)
            {
                message.Append(Duplicates == 1
                    ? " 1 was already in your vault."
                    : $" {Duplicates} were already in your vault.");
            }

            if (Cancelled && Processed < Total)
            {
                message.Append($" Cancelled; {Total - Processed} not processed.");
            }

            if (Failed > 0)
            {
                message.Append("\n\n").Append(Failed == 1 ? "1 failed:" : $"{Failed} failed:");
                foreach (KeyValuePair<string, string> failure in _failures.Take(MaxListedFailures))
                {
                    message.Append("\n• ").Append(failure.Key).Append(": ").Append(failure.Value);
                }

                if (Failed > MaxListedFailures)
                {
                    message.Append($"\n…and {Failed - MaxListedFailures} more.");
                }

                message.Append("\n\nDetails are in the SlingMD log.");
            }

            return message.ToString();
        }
    }
}
