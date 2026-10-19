using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace NoppaGovernanceApp
{
    public class MainForm : Form
    {
        private readonly EmbeddedWebServer _server;
        private WebView2? _webView;
        private CoreWebView2Environment? _environment;

        public MainForm(EmbeddedWebServer server)
        {
            _server = server;
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            Text = "Noppa | SharePoint & Teams Access Governance";
            Width = 1440;
            Height = 900;
            MinimumSize = new Size(1024, 700);
            StartPosition = FormStartPosition.CenterScreen;
            BackColor = Color.FromArgb(11, 16, 27); // Dark theme background
            ForeColor = Color.White;

            // Set Form Icon if app.ico exists
            string icoPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "app.ico");
            if (File.Exists(icoPath))
            {
                try { Icon = new Icon(icoPath); } catch { }
            }

            _webView = new WebView2
            {
                Dock = DockStyle.Fill
            };

            Controls.Add(_webView);

            Shown += async (s, e) => await InitializeWebViewAsync();
            FormClosing += (s, e) =>
            {
                _server.Dispose();
            };
        }

        private async Task InitializeWebViewAsync()
        {
            if (_webView == null) return;

            try
            {
                // Persistent user data folder in %LOCALAPPDATA%
                string userDataFolder = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "NoppaGovernance",
                    "WebView2Data"
                );
                Directory.CreateDirectory(userDataFolder);

                _environment = await CoreWebView2Environment.CreateAsync(null, userDataFolder);
                await _webView.EnsureCoreWebView2Async(_environment);

                // Configure WebView2 settings
                var settings = _webView.CoreWebView2.Settings;
                settings.IsStatusBarEnabled = false;
                settings.AreDefaultContextMenusEnabled = true;
                settings.IsZoomControlEnabled = true;
                settings.AreDevToolsEnabled = true;

                // Handle window.open (MSAL authentication popup & external links)
                _webView.CoreWebView2.NewWindowRequested += OnNewWindowRequested;

                // Navigate to local embedded web server
                _webView.CoreWebView2.Navigate(_server.BaseUrl);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Fout bij het laden van WebView2 runtime:\n\n{ex.Message}\n\nControleer of Microsoft Edge WebView2 Runtime is geïnstalleerd.",
                    "Fout bij opstarten",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }

        private void OnNewWindowRequested(object? sender, CoreWebView2NewWindowRequestedEventArgs e)
        {
            string uri = e.Uri ?? string.Empty;

            // If it's an external link (not localhost / not MS login), open in default browser
            bool isAuthUrl = uri.Contains("login.microsoftonline.com", StringComparison.OrdinalIgnoreCase) ||
                             uri.Contains("microsoftonline.com", StringComparison.OrdinalIgnoreCase) ||
                             uri.Contains("login.live.com", StringComparison.OrdinalIgnoreCase) ||
                             uri.StartsWith(_server.BaseUrl, StringComparison.OrdinalIgnoreCase);

            if (!isAuthUrl && (uri.StartsWith("http://", StringComparison.OrdinalIgnoreCase) || uri.StartsWith("https://", StringComparison.OrdinalIgnoreCase)))
            {
                e.Handled = true;
                try
                {
                    Process.Start(new ProcessStartInfo(uri) { UseShellExecute = true });
                }
                catch { }
                return;
            }

            // For MSAL Auth Popup, open a dedicated native popup window sharing the same environment
            var deferral = e.GetDeferral();

            var popupForm = new Form
            {
                Text = "Microsoft 365 Authenticatie - Noppa Governance",
                Width = 560,
                Height = 700,
                StartPosition = FormStartPosition.CenterParent,
                Icon = this.Icon,
                ShowInTaskbar = true,
                MinimizeBox = false,
                MaximizeBox = false,
                FormBorderStyle = FormBorderStyle.FixedDialog
            };

            var popupWebView = new WebView2
            {
                Dock = DockStyle.Fill
            };

            popupForm.Controls.Add(popupWebView);

            popupWebView.EnsureCoreWebView2Async(_environment).ContinueWith(t =>
            {
                if (popupForm.IsDisposed) return;

                popupForm.Invoke(() =>
                {
                    if (t.IsFaulted)
                    {
                        deferral.Complete();
                        popupForm.Close();
                        return;
                    }

                    e.NewWindow = popupWebView.CoreWebView2;
                    e.Handled = true;
                    deferral.Complete();

                    // When MSAL closes the popup after token acquisition
                    popupWebView.CoreWebView2.WindowCloseRequested += (s2, e2) =>
                    {
                        popupForm.Invoke(() => popupForm.Close());
                    };

                    popupForm.Show(this);
                });
            });
        }
    }
}
