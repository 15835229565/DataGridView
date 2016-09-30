using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;

namespace XLStoDataGridView
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFile = new OpenFileDialog();

            if (openFile.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                this.textBox_path.Text = openFile.FileName;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string pathConnection = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + textBox_path.Text + ";Extended Properties=\"Excel 12.0 Xml;HDR=YES;\";";
            OleDbConnection connection = new OleDbConnection(pathConnection);

            OleDbDataAdapter myDataAdapter = new OleDbDataAdapter("Select * from [" + textBox_sheet.Text + "$]", connection);
            DataTable dt = new DataTable();

            myDataAdapter.Fill(dt);

            dataGridView1.DataSource = dt;

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
