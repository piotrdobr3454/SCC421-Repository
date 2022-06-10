using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.IO;
using Newtonsoft.Json;
using Microsoft.Office.Core;
using PPT = Microsoft.Office.Interop.PowerPoint;

namespace MigrationFormApp
{
    public partial class Form1 : Form
    {
        //[DllImport("Gdi32.dll", EntryPoint = "CreateRoundRectRgn")]

        public Form1()
        {
            InitializeComponent();
            plnnav.Height = PPT2010to2016.Height;
            plnnav.Top = PPT2010to2016.Top;
            plnnav.Left = PPT2010to2016.Left;
            plnnav.BackColor = Color.FromArgb(46, 51, 73);
        }

        public void loadform(object Form)
        {
            if(this.MainPanel.Controls.Count > 0)
            {
                this.MainPanel.Controls.RemoveAt(0);
            }
            Form f = Form as Form;
            f.TopLevel = false;
            f.Dock = DockStyle.Fill;
            this.MainPanel.Controls.Add(f);
            this.MainPanel.Tag = f;
            f.Show();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void PPT2003to2010_Click(object sender, EventArgs e)
        {
            //plnnav.Height = PPT2003to2010.Height;
            //plnnav.Top = PPT2003to2010.Top;
            //plnnav.Left = PPT2003to2010.Left;
            //PPT2003to2010.BackColor = Color.FromArgb(46, 51, 73);
        }
        private void PPT2003to2010_Leave(object sender, EventArgs e)
        {
            //PPT2003to2010.BackColor = Color.FromArgb(0, 0, 64);
        }

        private void PPT2010to2016_Click(object sender, EventArgs e)
        {
            //plnnav.Height = PPT2010to2016.Height;
            //plnnav.Top = PPT2010to2016.Top;
            //plnnav.Left = PPT2010to2016.Left;
            //PPT2010to2016.BackColor = Color.FromArgb(46, 51, 73);
        }
        private void PPT2010to2016_Leave(object sender, EventArgs e)
        {
            //PPT2010to2016.BackColor = Color.FromArgb(0, 0, 64);
        }

        private void PPT2003to2016_Click(object sender, EventArgs e)
        {
            plnnav.Height = PPT2003to2016.Height;
            plnnav.Top = PPT2003to2016.Top;
            plnnav.Left = PPT2003to2016.Left;
            PPT2003to2016.BackColor = Color.FromArgb(46, 51, 73);
            loadform(new PPT2003_2016());
        }
        private void PPT2003to2016_Leave(object sender, EventArgs e)
        {
            PPT2003to2016.BackColor = Color.FromArgb(0, 0, 64);
        }

        private void PPT2013to2016_Click(object sender, EventArgs e)
        {
            //plnnav.Height = PPT2013to2016.Height;
            //plnnav.Top = PPT2013to2016.Top;
            //plnnav.Left = PPT2013to2016.Left;
            //PPT2013to2016.BackColor = Color.FromArgb(46, 51, 73);
        }
        private void PPT2013to2016_Leave(object sender, EventArgs e)
        {
            //PPT2013to2016.BackColor = Color.FromArgb(0, 0, 64);
        }
        private void Settings_Click(object sender, EventArgs e)
        {
            plnnav.Height = Settings.Height;
            plnnav.Top = Settings.Top;
            plnnav.Left = Settings.Left;
            Settings.BackColor = Color.FromArgb(46, 51, 73);
            loadform(new Settings());
        }
        private void Settings_Leave(object sender, EventArgs e)
        {
            Settings.BackColor = Color.FromArgb(0, 0, 64);
        }
        private void Xbutton_Click(object sender, EventArgs e)
        {
            this.Close();
        }

    }
}
