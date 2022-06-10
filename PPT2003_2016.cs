using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MigrationFormApp
{
    public partial class PPT2003_2016 : Form
    {
        public PPT2003_2016()
        {
            InitializeComponent();
        }

        private void Label1_Click(object sender, EventArgs e)
        {

        }

        private void Browse_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "PowerPoint Files|*.ppt";
            DialogResult result = openFileDialog1.ShowDialog(); // Show the dialog.
            if (result == DialogResult.OK) // Test result.
            {
                try
                {
                    string path = Path.GetDirectoryName(openFileDialog1.FileName);
                    string name = System.IO.Path.GetFileName(openFileDialog1.FileName);
                    label2.Text = path + "\\" + name;

                }
                catch (IOException)
                {
                }
            }
        }

        private String getLabel2Text()
        {
            return label2.Text;
        }
        private void ConvertButton_Click_1(object sender, EventArgs e)
        {
            string pathToPPT = label2.Text;
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "PowerPoint Files|*.ppt";

            DialogResult result = sfd.ShowDialog();
            if (result == DialogResult.OK) // Test result.
            {
                try
                {
                    string path = Path.GetDirectoryName(sfd.FileName);
                    string name = Path.GetFileName(sfd.FileName);
                    string exportPath = path + "\\" + name;
                    exportPath = exportPath.Substring(0, exportPath.IndexOf("."));
                    labWait.Visible = true;
                    if (ProgramConvertJSOn.CreateJSON(pathToPPT, exportPath))
                    {
                        if (Program.SaveData(exportPath + ".json"))
                        {
                            labWait.Text = "Conversion finished";
                        }
                    }
                }
                catch (IOException)
                {
                    labWait.Visible = true;
                    labWait.Text = "Unexpected error, please try again";
                }
            }
        }
    }
}
