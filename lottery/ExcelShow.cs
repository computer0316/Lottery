using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Lottery
{
    public partial class ExcelShow : Form
    {
        public ExcelShow()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Excel文件|*.xls";
            if(ofd.ShowDialog() == DialogResult.OK)
            {
                DataSet ds = Excel.OpenExcelUseOledb(ofd.FileName);
                dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;
                dataGridView1.DataSource = ds.Tables[0];
            }
        }
    }
}
