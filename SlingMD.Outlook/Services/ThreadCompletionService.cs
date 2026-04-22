using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Outlook;
using SlingMD.Outlook.Models;
using SlingMD.Outlook.Services.Formatting;

namespace SlingMD.Outlook.Services
{
    /// <summary>
    /// Finds emails in a conversation that have not yet been slung to the vault,
    /// enabling the "Complete Thread" feature.
    /// </summary>
    public class ThreadCompletionService
    {
        private readonly FileService _fileService;
        private readonly ObsidianSettings _settings;
        private readonly FrontmatterReader _frontmatter;

        public ThreadCompletionService(FileService fileService, ObsidianSettings settings)
        {
            _fileService = fileService;
            _settings = settings;
            _frontmatter = new FrontmatterReader();
        }

        /// <summary>
        /// Finds emails in a conversation that have not yet been slung to the vault.
        /// Searches the provided folder and the Sent Items folder for all emails matching
        /// the given conversationId, then filters out those already in the vault.
        /// </summary>
        /// <param name="conversationId">The conversation ID computed by ThreadService.GetConversationId.</param>
        /// <param name="searchFolder">The folder to search (typically Inbox).</param>
        /// <returns>List of MailItems not yet present in the vault.</returns>
        public List<MailItem> FindMissingEmails(string conversationId, MAPIFolder searchFolder)
        {
            HashSet<string> existingEntryIds = GetExistingEntryIds(conversationId);
            List<MailItem> missingEmails = new List<MailItem>();

            CollectMissingFromFolder(searchFolder, existingEntryIds, missingEmails);

            try
            {
                MAPIFolder sentItems = searchFolder.Session.GetDefaultFolder(OlDefaultFolders.olFolderSentMail);
                CollectMissingFromFolder(sentItems, existingEntryIds, missingEmails);
            }
            catch (System.Exception)
            {
                // If Sent Items is inaccessible, skip it silently
            }

            return missingEmails;
        }

        /// <summary>
        /// Collects emails from a folder whose ConversationTopic matches and whose EntryID
        /// is not already present in the vault.
        /// </summary>
        private void CollectMissingFromFolder(MAPIFolder folder, HashSet<string> existingEntryIds, List<MailItem> results)
        {
            if (folder == null)
            {
                return;
            }

            try
            {
                Items items = folder.Items;
                foreach (object item in items)
                {
                    MailItem mail = item as MailItem;
                    if (mail == null)
                    {
                        continue;
                    }

                    string entryId = mail.EntryID;
                    if (!string.IsNullOrEmpty(entryId) && !existingEntryIds.Contains(entryId))
                    {
                        results.Add(mail);
                    }
                }
            }
            catch (System.Exception)
            {
                // Skip folders that cannot be enumerated
            }
        }

        /// <summary>
        /// Scans all .md files in the vault inbox path (recursively) and returns the set
        /// of entryId values found in notes whose threadId matches the given conversationId.
        /// </summary>
        internal HashSet<string> GetExistingEntryIds(string conversationId)
        {
            HashSet<string> entryIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            string inboxPath = _fileService.GetInboxPath();
            if (!Directory.Exists(inboxPath))
            {
                return entryIds;
            }

            string[] files = Directory.GetFiles(inboxPath, "*.md", SearchOption.AllDirectories);
            foreach (string file in files)
            {
                try
                {
                    string content = File.ReadAllText(file);
                    string fileThreadId = _frontmatter.ExtractThreadId(content);
                    if (fileThreadId != conversationId)
                    {
                        continue;
                    }

                    string entryId = _frontmatter.ExtractEntryId(content);
                    if (!string.IsNullOrEmpty(entryId))
                    {
                        entryIds.Add(entryId);
                    }
                }
                catch (System.Exception)
                {
                    // Skip files that cannot be read
                }
            }

            return entryIds;
        }
    }
}
