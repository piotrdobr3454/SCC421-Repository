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
    public partial class Settings : Form
    {
        public Settings()
        {
            InitializeComponent();
        }

        private void Label1_Click(object sender, EventArgs e)
        {

        }

        private void Browse_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "Excel Files|*.xlsx";
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

        private void button1_Click(object sender, EventArgs e)
        {
            string pathToPPT = label2.Text;
            if (pathToPPT == "Name of file")
            {
                MessageBox.Show("Error, library needs to be set");
            }
            else if(pathToPPT.Contains("xlsx"))
            {
                Engine.setPathToExcel(pathToPPT);
            }
            else
            {
                MessageBox.Show("Improper format of Excel file");
            }
        }
    }
}
