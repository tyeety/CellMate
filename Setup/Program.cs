using System;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CellMate.Setup
{
    internal static class Program
    {
        [STAThread]
        private static int Main(string[] args)
        {
            if (args.Contains("--self-test"))
            {
                try { Console.WriteLine(InstallerTests.Run()); return 0; }
                catch (Exception ex) { Console.Error.WriteLine(ex); return 1; }
            }
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            using (var mutex = new Mutex(false, @"Local\CellMate.Setup." + System.Security.Principal.WindowsIdentity.GetCurrent().User.Value))
            {
                bool acquired;
                try { acquired = mutex.WaitOne(0); } catch (AbandonedMutexException) { acquired = true; }
                if (!acquired) { MessageBox.Show("CellMate Setup is already open.", "CellMate", MessageBoxButtons.OK, MessageBoxIcon.Information); return 2; }
                try { Application.Run(new SetupForm(args.Contains("--uninstall"), args.Contains("--preview"))); }
                finally { mutex.ReleaseMutex(); }
            }
            return 0;
        }
    }

    internal sealed class SetupForm : Form
    {
        private readonly Installer installer = new Installer();
        private readonly bool uninstall;
        private readonly bool preview;
        private ExcelInstallation excel;
        private readonly Label title;
        private readonly Label description;
        private readonly Label status;
        private readonly Label statusDetail;
        private readonly Label pathLabel;
        private readonly BrandButton action;
        private readonly BrandButton cancel;
        private readonly ProgressBar progress;
        private readonly System.Windows.Forms.Timer timer;
        private bool working, waiting, done;

        public SetupForm(bool uninstall, bool preview)
        {
            this.uninstall = uninstall;
            this.preview = preview;
            Theme.StyleWindow(this);
            Text = "CellMate Setup" + (preview ? " - Preview" : "");
            ClientSize = new Size(920, 608);
            Controls.Add(new BrandRail { Dock = DockStyle.Left, Width = 290 });
            var content = new Panel { Bounds = new Rectangle(290, 0, 630, 608), BackColor = Theme.Background };
            Controls.Add(content);
            content.Controls.Add(Place(Theme.Label(uninstall ? "CELLMATE / UNINSTALL" : "CELLMATE FOR EXCEL", 9, true, Theme.Muted), 36, 28, 550, 24));
            title = Place(Theme.Label(uninstall ? "Time to say goodbye?" : "Meet your new\nExcel companion.", 28, true), 34, 66, 550, 108);
            content.Controls.Add(title);
            description = Place(Theme.Label(uninstall ? "Remove the add-in. Your personal shortcuts will stay\nright here, ready if you return." : "Less repetition. More momentum.\nGet CellMate up and running in a single click.", 11, false, Theme.Body), 36, 184, 550, 57);
            content.Controls.Add(description);
            var features = Place(Theme.Label("Custom shortcuts     /     Smarter rows     /     A smoother flow", 9, true, Theme.Muted), 36, 258, 550, 28);
            content.Controls.Add(features);
            var card = new Card { Bounds = new Rectangle(34, 304, 556, 112) };
            status = Place(Theme.Label("Checking your Excel installation...", 11, true), 22, 16, 510, 30);
            statusDetail = Place(Theme.Label("Just a moment.", 9, false, Theme.Body), 22, 51, 510, 42);
            card.Controls.AddRange(new Control[] { status, statusDetail });
            content.Controls.Add(card);
            pathLabel = Place(Theme.Label("Installs for you. No administrator access needed.", 9, false, Theme.Body), 36, 428, 550, 26);
            content.Controls.Add(pathLabel);
            progress = new ProgressBar { Bounds = new Rectangle(36, 465, 550, 5), Style = ProgressBarStyle.Marquee, Visible = false, MarqueeAnimationSpeed = 22 };
            content.Controls.Add(progress);
            action = new BrandButton { Text = uninstall ? "Uninstall CellMate" : "Install CellMate", Primary = true, Bounds = new Rectangle(360, 492, 228, 48), Enabled = false };
            action.Click += async (_, __) => await Act();
            cancel = new BrandButton { Text = "Cancel", Bounds = new Rectangle(236, 492, 110, 48), DialogResult = DialogResult.Cancel };
            content.Controls.AddRange(new Control[] { action, cancel });
            content.Controls.Add(Place(Theme.Label("CellMate " + Theme.Version + "    /    Windows + Excel 2016 or later", 9, false, Theme.Muted), 36, 562, 550, 24));
            AcceptButton = action;
            CancelButton = cancel;
            timer = new System.Windows.Forms.Timer { Interval = 900 };
            timer.Tick += async (_, __) => { if (waiting && !Installer.IsExcelRunning()) { waiting = false; timer.Stop(); await Execute(); } };
            Shown += (_, __) => DetectExcel();
            FormClosing += (_, e) => { if (working) e.Cancel = true; };
        }
        private static T Place<T>(T control, int x, int y, int w, int h) where T : Control { control.Bounds = new Rectangle(x, y, w, h); return control; }
        private void DetectExcel()
        {
            try
            {
                excel = ExcelInstallation.Find();
                if (excel == null && !uninstall)
                {
                    status.Text = "Excel desktop was not found";
                    statusDetail.Text = "Install Excel 2016 or later for Windows, then reopen this installer.";
                    action.Text = "Check again";
                }
                else
                {
                    status.Text = uninstall ? "Your shortcut settings will be kept" : excel.Description + " detected";
                    statusDetail.Text = Installer.IsExcelRunning() ? "Excel is open. Save your work and close it before continuing." : uninstall ? "Your workbooks and other Excel add-ins stay untouched." : "The right add-in version will be selected automatically.\nYour existing shortcut settings will be kept.";
                }
                action.Enabled = true;
            }
            catch (Exception ex) { ShowFailure(ex); }
        }
        private async Task Act()
        {
            if (done)
            {
                if (!uninstall && !preview && excel != null) Process.Start(new ProcessStartInfo(excel.Path) { UseShellExecute = true });
                Close(); return;
            }
            if (excel == null && !uninstall) { DetectExcel(); return; }
            if (!preview && Installer.IsExcelRunning())
            {
                waiting = true;
                title.Text = "One quick pause.";
                description.Text = "Save your work and close Excel.\nWe will continue automatically when it closes.";
                status.Text = "Waiting for Excel to close";
                statusDetail.Text = "CellMate will never close your workbooks for you.\nYou can cancel at any time.";
                action.Text = "Waiting for Excel...";
                action.Enabled = false;
                timer.Start();
                return;
            }
            await Execute();
        }
        private async Task Execute()
        {
            working = true;
            action.Enabled = cancel.Enabled = false;
            progress.Visible = true;
            title.Text = uninstall ? "Removing CellMate..." : "A better flow\nis on its way.";
            status.Text = uninstall ? "Removing the add-in" : "Installing CellMate";
            statusDetail.Text = "This should only take a moment.";
            try
            {
                if (!preview) await Task.Run(() => { if (uninstall) installer.Uninstall(); else installer.Install(excel); });
                done = true;
                title.Text = uninstall ? "All taken care of." : "You're all set.";
                description.Text = uninstall ? "CellMate has been removed from Excel.\nYour personal shortcut settings are safe." : "CellMate is ready to make everyday work easier.\nOpen Excel and find the CellMate tab.";
                status.Text = uninstall ? "Uninstalled successfully" : "Installed successfully";
                statusDetail.Text = uninstall ? "You can reinstall CellMate whenever you like." : "Ctrl+D duplicates rows. Ctrl+Alt+K opens settings.\nEvery shortcut is yours to change.";
                pathLabel.Text = uninstall ? "Your Excel workbooks have not been changed." : "Manage or remove CellMate from Windows Settings > Apps.";
                action.Text = uninstall || preview ? "Done" : "Open Excel";
                cancel.Text = "Close";
                if (uninstall && !preview) installer.DeleteUninstallerAfterExit();
            }
            catch (Exception ex) { ShowFailure(ex); }
            finally { working = false; progress.Visible = false; action.Enabled = cancel.Enabled = true; }
        }
        private void ShowFailure(Exception ex)
        {
            title.Text = "Let's try that again.";
            description.Text = "CellMate could not complete the operation.\nReview the details below, then try again.";
            status.Text = "Setup needs your attention";
            status.ForeColor = Theme.Error;
            statusDetail.Text = ex.GetBaseException().Message;
            action.Text = "Try again";
            action.Enabled = true;
        }
        protected override void Dispose(bool disposing) { if (disposing) timer?.Dispose(); base.Dispose(disposing); }
    }
}
