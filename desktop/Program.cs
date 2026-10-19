using System;
using System.Windows.Forms;

namespace NoppaGovernanceApp
{
    internal static class Program
    {
        [STAThread]
        static void Main()
        {
            ApplicationConfiguration.Initialize();

            EmbeddedWebServer? server = null;
            try
            {
                server = new EmbeddedWebServer(3334);
                server.Start();

                Application.Run(new MainForm(server));
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Er is een onverwachte fout opgetreden bij het starten van de applicatie:\n\n{ex.Message}",
                    "Noppa Governance Fout",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
            finally
            {
                server?.Dispose();
            }
        }
    }
}