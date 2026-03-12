using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Windows.Forms;

namespace SlingMD.Outlook.Services
{
    internal static class SupportService
    {
        internal const string BuyMeACoffeeUrl = "https://buymeacoffee.com/plainsprepper";
        internal const string SupportWindowTitle = "Support SlingMD";
        internal const string SupportHeadline = "\u2615 Like what I'm building? Help fuel my next project (or my next coffee)!";
        internal const string SupportSubtitle = "Support me on Buy Me a Coffee \uD83D\uDCBB\uD83E\uDDF5\uD83D\uDD25";

        internal static string GetSupportMessage()
        {
            return SupportHeadline + Environment.NewLine + SupportSubtitle;
        }

        internal static void ShowBuyMeACoffeePrompt(IWin32Window owner = null)
        {
            string promptMessage = GetSupportMessage()
                + Environment.NewLine
                + Environment.NewLine
                + BuyMeACoffeeUrl
                + Environment.NewLine
                + Environment.NewLine
                + "Would you like to open the link now?";

            DialogResult result = MessageBox.Show(
                owner,
                promptMessage,
                SupportWindowTitle,
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Information);

            if (result == DialogResult.Yes)
            {
                OpenBuyMeACoffeeLink(owner);
            }
        }

        internal static void OpenBuyMeACoffeeLink(IWin32Window owner = null)
        {
            try
            {
                ProcessStartInfo startInfo = new ProcessStartInfo
                {
                    FileName = BuyMeACoffeeUrl,
                    UseShellExecute = true
                };

                Process.Start(startInfo);
            }
            catch (Win32Exception ex)
            {
                ShowUnableToOpenLinkMessage(owner, ex.Message);
            }
            catch (InvalidOperationException ex)
            {
                ShowUnableToOpenLinkMessage(owner, ex.Message);
            }
        }

        private static void ShowUnableToOpenLinkMessage(IWin32Window owner, string errorMessage)
        {
            MessageBox.Show(
                owner,
                "Unable to open the Buy Me a Coffee link."
                    + Environment.NewLine
                    + Environment.NewLine
                    + errorMessage,
                SupportWindowTitle,
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }
    }
}
