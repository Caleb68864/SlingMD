/*
 * BASELINE CAPTURE PROCEDURE (run once, then permanently skip):
 *
 *  1. Check out current main HEAD's ContactService.cs (pre-SS-09).
 *  2. Run the CaptureBaseline_RunOnceManually test below (remove [Fact(Skip=...)] temporarily).
 *     It writes expected-outputs.json to Fixtures/ToggleOffBaseline/.
 *  3. Commit that JSON alongside this file.
 *  4. Re-add [Fact(Skip = "Capture only — run once then skip permanently")] and commit.
 *
 * The JSON files in Fixtures/ToggleOffBaseline/ are IMMUTABLE for the duration of this feature.
 * Any intentional behavior change requires a dedicated spec escalation before modifying them.
 */

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using SlingMD.Outlook.Models;
using SlingMD.Outlook.Services;
using Xunit;

namespace SlingMD.Tests.Services
{
    public class ContactMatcherToggleOffRegressionTests : IDisposable
    {
        private static readonly string FixturesDir = Path.Combine(
            AppDomain.CurrentDomain.BaseDirectory, "Fixtures", "ToggleOffBaseline");

        private readonly string _baseDir;

        public ContactMatcherToggleOffRegressionTests()
        {
            _baseDir = Path.Combine(Path.GetTempPath(), "SlingMDTests",
                "ToggleOff_" + Guid.NewGuid().ToString("N").Substring(0, 8));
            Directory.CreateDirectory(_baseDir);
        }

        public void Dispose()
        {
            try
            {
                if (Directory.Exists(_baseDir))
                {
                    Directory.Delete(_baseDir, true);
                }
            }
            catch (System.Exception)
            {
                // Best-effort cleanup
            }
        }

        // ---------------------------------------------------------------------------
        // Theory data loader
        // ---------------------------------------------------------------------------

        public static IEnumerable<object[]> LoadTestCases()
        {
            string inputsPath = Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory, "Fixtures", "ToggleOffBaseline", "inputs.json");
            string expectedPath = Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory, "Fixtures", "ToggleOffBaseline", "expected-outputs.json");

            JObject inputs = JObject.Parse(File.ReadAllText(inputsPath));
            JObject expectedOutputs = JObject.Parse(File.ReadAllText(expectedPath));

            JArray inputCases = (JArray)inputs["cases"];
            JArray expectedCases = (JArray)expectedOutputs["cases"];

            Dictionary<string, JToken> expectedByName = expectedCases
                .ToDictionary(c => c["name"].Value<string>(), c => c["expected"]);

            foreach (JToken inputCase in inputCases)
            {
                string name = inputCase["name"].Value<string>();
                JToken input = inputCase["input"];
                JToken expected = expectedByName[name];
                yield return new object[] { name, input.ToString(), expected.ToString() };
            }
        }

        // ---------------------------------------------------------------------------
        // Main regression theory
        // ---------------------------------------------------------------------------

        [Theory]
        [MemberData(nameof(LoadTestCases))]
        public void ToggleOff_ReproducesBaselineBehavior(string caseName, string inputJson, string expectedJson)
        {
            JObject input = JObject.Parse(inputJson);
            JObject expected = JObject.Parse(expectedJson);

            string displayName = input["displayName"].Value<string>();
            string email = input["email"].Value<string>();
            JArray vaultLayout = (JArray)(input["vaultLayout"] ?? new JArray());

            string vaultBase = Path.Combine(_baseDir, caseName);
            string vaultName = "TestVault";
            string contactsFolder = "Contacts";
            string fullVaultPath = Path.Combine(vaultBase, vaultName);
            string contactsPath = Path.Combine(fullVaultPath, contactsFolder);
            Directory.CreateDirectory(contactsPath);

            // Materialise vault layout files
            foreach (JToken entry in vaultLayout)
            {
                string relativePath = entry["path"].Value<string>().Replace('/', Path.DirectorySeparatorChar);
                string filePath = Path.Combine(fullVaultPath, relativePath);
                Directory.CreateDirectory(Path.GetDirectoryName(filePath));
                File.WriteAllText(filePath, "# " + displayName);
            }

            ObsidianSettings settings = new ObsidianSettings
            {
                VaultBasePath = vaultBase,
                VaultName = vaultName,
                ContactsFolder = contactsFolder,
                EnableContactSaving = true,
                SearchEntireVaultForContacts = false,
                EnableContactFuzzyMatching = false,
                ContactLinkFormat = "[[{FullName}]]",
                ContactFilenameFormat = "{ContactName}"
            };

            FileService fileService = new FileService(settings);
            TemplateService templateService = new TemplateService(fileService);
            ContactService service = new ContactService(fileService, templateService);

            // --- contactExists ---
            bool actualContactExists = service.ContactExists(displayName);
            Assert.Equal(expected["contactExists"].Value<bool>(), actualContactExists);

            // --- formatContactLink (private — accessed via reflection) ---
            string actualFormatContactLink = InvokeFormatContactLink(service, displayName, email);
            Assert.Equal(expected["formatContactLink"].Value<string>(), actualFormatContactLink);

            // --- buildLinkedNames: mirrors FormatContactLink per recipient ---
            string[] actualBuildLinkedNames = new[] { actualFormatContactLink };
            string[] expectedBuildLinkedNames = expected["buildLinkedNames"]
                .ToObject<string[]>();
            Assert.Equal(expectedBuildLinkedNames, actualBuildLinkedNames);

            // --- buildContactFileName (private — accessed via reflection) ---
            string actualBuildContactFileName = InvokeBuildContactFileName(service, displayName);
            Assert.Equal(expected["buildContactFileName"].Value<string>(), actualBuildContactFileName);

            // --- createContactNotePath ---
            string actualNotePath = service.GetManagedContactNotePath(displayName);
            string normalizedActual = actualNotePath.Replace('\\', '/');
            string expectedSuffix = expected["createContactNotePath"].Value<string>();
            Assert.True(
                normalizedActual.EndsWith(expectedSuffix, StringComparison.OrdinalIgnoreCase),
                $"Case '{caseName}': expected path ending '{expectedSuffix}' but got '{normalizedActual}'");
        }

        // ---------------------------------------------------------------------------
        // Capture helper (skip permanently after first run)
        // ---------------------------------------------------------------------------

        [Fact(Skip = "Capture only — run once against main HEAD then skip permanently")]
        public void CaptureBaseline_RunOnceManually()
        {
            string inputsPath = Path.Combine(FixturesDir, "inputs.json");
            JObject inputs = JObject.Parse(File.ReadAllText(inputsPath));
            JArray inputCases = (JArray)inputs["cases"];

            string vaultName = "TestVault";
            string contactsFolder = "Contacts";

            JArray capturedCases = new JArray();

            foreach (JToken inputCase in inputCases)
            {
                string name = inputCase["name"].Value<string>();
                JToken caseInput = inputCase["input"];
                string displayName = caseInput["displayName"].Value<string>();
                string email = caseInput["email"].Value<string>();
                JArray vaultLayout = (JArray)(caseInput["vaultLayout"] ?? new JArray());

                string caseDir = Path.Combine(_baseDir, "capture_" + name);
                string fullVaultPath = Path.Combine(caseDir, vaultName);
                string contactsPath = Path.Combine(fullVaultPath, contactsFolder);
                Directory.CreateDirectory(contactsPath);

                foreach (JToken entry in vaultLayout)
                {
                    string relativePath = entry["path"].Value<string>().Replace('/', Path.DirectorySeparatorChar);
                    string filePath = Path.Combine(fullVaultPath, relativePath);
                    Directory.CreateDirectory(Path.GetDirectoryName(filePath));
                    File.WriteAllText(filePath, "# " + displayName);
                }

                ObsidianSettings settings = new ObsidianSettings
                {
                    VaultBasePath = caseDir,
                    VaultName = vaultName,
                    ContactsFolder = contactsFolder,
                    EnableContactSaving = true,
                    SearchEntireVaultForContacts = false,
                    EnableContactFuzzyMatching = false,
                    ContactLinkFormat = "[[{FullName}]]",
                    ContactFilenameFormat = "{ContactName}"
                };

                FileService fileService = new FileService(settings);
                TemplateService templateService = new TemplateService(fileService);
                ContactService service = new ContactService(fileService, templateService);

                bool contactExists = service.ContactExists(displayName);
                string formatContactLink = InvokeFormatContactLink(service, displayName, email);
                string buildContactFileName = InvokeBuildContactFileName(service, displayName);
                string notePath = service.GetManagedContactNotePath(displayName)
                    .Replace(fullVaultPath.Replace('\\', '/'), string.Empty)
                    .Replace(fullVaultPath, string.Empty)
                    .TrimStart('/', '\\')
                    .Replace('\\', '/');

                capturedCases.Add(new JObject
                {
                    ["name"] = name,
                    ["expected"] = new JObject
                    {
                        ["contactExists"] = contactExists,
                        ["formatContactLink"] = formatContactLink,
                        ["buildLinkedNames"] = new JArray(formatContactLink),
                        ["buildContactFileName"] = buildContactFileName,
                        ["createContactNotePath"] = notePath
                    }
                });
            }

            JObject output = new JObject
            {
                ["_comment"] = "Baseline captured from main HEAD ContactService (pre-SS-09). DO NOT modify without a dedicated spec escalation.",
                ["cases"] = capturedCases
            };

            File.WriteAllText(Path.Combine(FixturesDir, "expected-outputs.json"),
                output.ToString(Newtonsoft.Json.Formatting.Indented));
        }

        // ---------------------------------------------------------------------------
        // Reflection helpers
        // ---------------------------------------------------------------------------

        private static string InvokeFormatContactLink(ContactService service, string displayName, string email)
        {
            MethodInfo method = typeof(ContactService).GetMethod(
                "FormatContactLink",
                BindingFlags.NonPublic | BindingFlags.Instance,
                null,
                new[] { typeof(string), typeof(string) },
                null);

            if (method == null)
            {
                throw new InvalidOperationException(
                    "ContactService.FormatContactLink(string, string) private method not found. " +
                    "If the method signature changed, update this test.");
            }

            return (string)method.Invoke(service, new object[] { displayName, email });
        }

        private static string InvokeBuildContactFileName(ContactService service, string contactName)
        {
            MethodInfo method = typeof(ContactService).GetMethod(
                "BuildContactFileName",
                BindingFlags.NonPublic | BindingFlags.Instance,
                null,
                new[] { typeof(string) },
                null);

            if (method == null)
            {
                throw new InvalidOperationException(
                    "ContactService.BuildContactFileName(string) private method not found. " +
                    "If the method signature changed, update this test.");
            }

            return (string)method.Invoke(service, new object[] { contactName });
        }
    }
}
