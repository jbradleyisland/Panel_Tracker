using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Panel_Tracker
{
    public partial class Tracker : Form
    {
        public bool filled { get; set; }
        public string bimLink { get; set; }
        public string sharepointLink { get; set; }
        public string namecolumn { get; set; }
        public string namerow { get; set; }
        public string statuscolumn { get; set; }
        public string statusrow { get; set; }
        public bool useParameter { get; set; }
        public string parameter { get; set; }
        public string elementType { get; set; }

        public const int WM_NCLBUTTONDOWN = 0xA1;
        public const int HT_CAPTION = 0x2;

        [DllImportAttribute("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        [DllImportAttribute("user32.dll")]
        public static extern bool ReleaseCapture();

        public bool exit = false;

        public Tracker()
        {
            InitializeComponent();

            foreach (string s in Panel_Tracker.Command.categories["Assemblies"])
            {
                this.eleComboBox1.Items.Add(s);
            }
            foreach (string p in Panel_Tracker.Command.assemParameters)
            {
                this.parameterCB.Items.Add(p);
            }
        }

        private void panel1_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }
        }

        private void submit_lbl_Click(object sender, EventArgs e)
        {
            if (this.bimDocsLink_txtbox.Text != "" && this.sharePointLink_txtbox.Text != "" && this.statusColumn_txtbox.Text != "" && this.statusRow_txtbox.Text != "")
            {
                filled = true;

                bimLink = this.bimDocsLink_txtbox.Text;
                sharepointLink = this.sharePointLink_txtbox.Text;
                namecolumn = this.panelNameCol_txtbox.Text;
                namerow = this.panelNameRow_txtbox.Text;
                statuscolumn = this.statusColumn_txtbox.Text;
                statusrow = this.statusRow_txtbox.Text;
            }

            this.Close();
            exit = true;
        }

        private void Tracker_Load(object sender, EventArgs e)
        {

        }

        private void parameterCBox_CheckedChanged(object sender, EventArgs e)
        {
            if (parameterCBox.Checked == true)
            {
                parameterCB.Enabled = true;
                useParameter = true;

            }
            else if (parameterCBox.Checked == false)
            {
                parameterCB.Enabled = false;
                useParameter = false;
            }
        }

        private void parameterCB_SelectedIndexChanged(object sender, EventArgs e)
        {
            parameter = parameterCB.SelectedItem.ToString();
        }

        private void eleComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            elementType = eleComboBox1.SelectedItem.ToString();

            this.parameterCB.Items.Clear();
            foreach (string p in Panel_Tracker.Command.assemParameters)
            {
                this.parameterCB.Items.Add(p);
            }
        }
    }
}
