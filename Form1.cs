using System;
using System.Diagnostics;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Net;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Runtime.InteropServices;
using System.Drawing;
using System.Data;
using System.Management;

/*
 * Author: Chris Kye Harris
 */

namespace CIRedux
{   
    public partial class CIRedux : Form
    {
        // NEEDED TO MAKE BORDERLESS FORM MOVABLE
        public const int WM_NCLBUTTONDOWN = 0xA1;
        public const int HT_CAPTION = 0x2;
        [DllImportAttribute("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        [DllImportAttribute("user32.dll")]
        public static extern bool ReleaseCapture();

        public CIRedux() => InitializeComponent();

        // THIS IS OUR APPLICATION CLOSE BUTTON
        private void CloseBTN_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void CloseBTN_MouseHover(object sender, EventArgs e)
        {
            this.closeBTN.ForeColor = Color.FromArgb(231, 76, 60);
        }

        private void CloseBTN_MouseLeave(object sender, EventArgs e)
        {
            this.closeBTN.ForeColor = Color.FromArgb(255, 255, 255);
        }

        private DateTime BootTime = new DateTime();

        protected override void OnShown(EventArgs e)
        {
            // GET IP FROM ADAPTER THAT CONTACTS 8.8.8.8
            string localIP = string.Empty;
            string errmsg1 = "Something went wrong while trying to connect to socket.\nCheck for issues with Network Adapter.\nRecommended to restart system.";
            string errtitle = "Computer Info";
            try
            {
                using (var socket = new Socket(AddressFamily.InterNetwork, SocketType.Dgram, 0))
                {
                    socket.ConnectAsync("8.8.8.8", 65530); // Attempts to connect to Google DNS via socket
                    localIP = (socket.LocalEndPoint as IPEndPoint).Address.ToString(); // Get local IP from Network Adapter 
                }
            }
            catch //(Exception ex)
            {
                MessageBox.Show(errmsg1, errtitle);
            }

            // GET OS INFO
            string HKLMWinNTCurrent = @"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion";
            string osName = get(() => Registry.GetValue(HKLMWinNTCurrent, "ProductName", "").ToString());
            string osRelease = get(() => Registry.GetValue(HKLMWinNTCurrent, "DisplayVersion", "").ToString());
            string Platform = HardwareInfo.GetOSInformation() + " " + osRelease;
            string get(Func<string> func)
            {
                try { return func(); }
                catch { return "(undefined)"; }
            }

            // VPN STATUS (LOOKS FOR GLOBAL PROTECT PANGP ADAPTER)
            NetworkInterface[] adapters = NetworkInterface.GetAllNetworkInterfaces();
            if (adapters.Where(x => x.Description.StartsWith("PANGP")).FirstOrDefault() != null)
            {
                vpnStatus.Text = "Connected!"; // Display VPN Status as Connected once connected to Global Protect VPN
                vpnStatus.ForeColor = ColorTranslator.FromHtml("#2ecc71"); // Set VPN Status text to Green on connect
            }
            else
            {
                vpnStatus.Text = "Disconnected!"; // Display VPN Status as Disconnected if not connected to Global Protect VPN
                vpnStatus.ForeColor = ColorTranslator.FromHtml("#d63031"); // Set VPN Status text to Red on connect
            }

            // Get the last time computer rebooted
            SelectQuery query = new SelectQuery("SELECT LastBootUpTime FROM Win32_OperatingSystem WHERE Primary = 'true'");
            ManagementObjectSearcher searcher = new ManagementObjectSearcher(query);
            foreach (ManagementObject mo in searcher.Get())
            {
                BootTime = ManagementDateTimeConverter.ToDateTime(mo.Properties["LastBootUpTime"].Value.ToString());
                uptimeLB.Text = BootTime.ToLongDateString(); // Displays Last boot date.
                uptimeLB.ForeColor = ColorTranslator.FromHtml("#f1c40f"); // Sets Text Color of "Last Boot"
            }

            //RCIT Image Version key
            string RCVKey = @"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Management\Publishers\RC";
            string VKey = get(() => Registry.GetValue(RCVKey, "Version", "").ToString());

            // INFO WE NEED TO KNOW
            ipTB.Text = localIP; // Display local IP from Network Adapter we use to get out to the internet.
            osTB.Text = Platform; // Display OS (ex. Microsoft Windows 10 Enterprise 21H2)
            domainTB.Text = System.Environment.UserDomainName; // Display domain (ex. RIVCOCA)
            compNameTB.Text = System.Environment.MachineName; // Display Computer Name (ex. ITD2UA7222CFR)
            compmodTB.Text = HardwareInfo.GetBoardMaker() + " " + HardwareInfo.GetBoardProductId(); // Display Computer Model (ex. HP 82C0)
            serialtagTB.Text = HardwareInfo.GetBIOSserNo(); // Display Serial/Service TAG (ex. Dell Service Tag/HP Serial Number)
            versionkeyTB.Text = VKey; // Displays Version of RCIT Image used for system.
        }
        // This button is used to refressh info that may change while application is open.
        private void refreshBTN_Click(object sender, EventArgs e)
        {
            // GET IP FROM ADAPTER THAT CONTACTS 8.8.8.8
            string localIP = string.Empty;
            string errmsg1 = "Something went wrong while trying to connect to socket.\nCheck for issues with Network Adapter.\nRecommended to restart system.";
            string errtitle = "Computer Info";
            try
            {
                using (var socket = new Socket(AddressFamily.InterNetwork, SocketType.Dgram, 0))
                {
                    socket.ConnectAsync("8.8.8.8", 65530); // Attempts to connect to Google DNS via socket
                    localIP = (socket.LocalEndPoint as IPEndPoint).Address.ToString(); // Get local IP from Network Adapter 
                }
            }
            catch //(Exception ex)
            {
                MessageBox.Show(errmsg1, errtitle);
            }

            // VPN STATUS (LOOKS FOR GLOBAL PROTECT PANGP ADAPTER)
            NetworkInterface[] adapters = NetworkInterface.GetAllNetworkInterfaces();
            if (adapters.Where(x => x.Description.StartsWith("PANGP")).FirstOrDefault() == null)
            {
                vpnStatus.Text = "Disconnected!"; // Display VPN Status as Disconnected if not connected to Global Protect VPN
                vpnStatus.ForeColor = ColorTranslator.FromHtml("#d63031"); // Set VPN Status text to Red on connect
            }
            else
            {
                vpnStatus.Text = "Connected!"; // Display VPN Status as Connected once connected to Global Protect VPN
                vpnStatus.ForeColor = ColorTranslator.FromHtml("#2ecc71"); // Set VPN Status text to Green on connect
            }

            ipTB.Text = localIP; // Display local IP from Network Adapter we use to get out to the internet.
        }

        // ImageList index value for the hover image.
        private void qsBTN_MouseHover(object sender, EventArgs e) => qsBTN.ImageIndex = 1;

        // ImageList index value for the normal image.
        private void qsBTN_MouseLeave(object sender, EventArgs e) => qsBTN.ImageIndex = 0;

        // Mouse down events for moving form.
        private void Form1_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
        }

        private void panel1_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
        }

        private void qsBTN_Click(object sender, EventArgs e)
        {
            string noqs = "Quick Assist does not exist on this system.";
            string errtitle = "Computer Info";
            try
            {
                Process process = new Process();
                ProcessStartInfo startInfo = new ProcessStartInfo();
                startInfo.WindowStyle = ProcessWindowStyle.Hidden;
                startInfo.FileName = "cmd.exe";
                startInfo.Arguments = @"/C explorer.exe shell:appsFolder\MicrosoftCorporationII.QuickAssist_8wekyb3d8bbwe!App";
                process.StartInfo = startInfo;
                process.Start();
            }
            catch
            {
                MessageBox.Show(noqs, errtitle);
            }
        }

        private void vpnStatus_Click(object sender, EventArgs e)
        {
            Process process = new Process();
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.WindowStyle = ProcessWindowStyle.Hidden;
            startInfo.FileName = "cmd.exe";
            startInfo.WorkingDirectory = @"C:\Program Files\Palo Alto Networks\GlobalProtect";
            startInfo.Arguments = "/C PanGPA.exe";
            process.StartInfo = startInfo;
            process.Start();
        }
    }
}
