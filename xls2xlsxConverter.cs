using System;
using System.Data;
using System.Windows.Forms;

namespace xls2xlsxConverter
{
    public partial class xls2xlsxConverter : Form
    {
        public xls2xlsxConverter()
        {
            InitializeComponent();
            Init_dGV();
        }

        public void Init_dGV()
        {
            DataTable dt = new DataTable();

            dt.Columns.Add("No.", typeof(int));
            dt.Columns.Add("FileName", typeof(string));
            dt.Columns.Add("Progress", typeof(string));

            dGV_FileList.DataSource = dt;
        }

        private void btn_folder_Click(object sender, EventArgs e)
        {
            using (var v_FBD = new FolderBrowserDialog())
            {
                DialogResult result = v_FBD.ShowDialog();

                if (result == DialogResult.OK)
                {
                    txt_selectedPath.Text = v_FBD.SelectedPath;
                }
            }
        }

        private void btn_FileAdd_Click(object sender, EventArgs e)
        {

        }

        private void btn_FileDelete_Click(object sender, EventArgs e)
        {
            if (dGV_FileList.SelectedRows.Count > 0)
            {
                DialogResult result = MessageBox.Show("Would you like to delete the selected row(s)", "Delete Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    foreach (DataGridViewRow row in dGV_FileList.SelectedRows)
                    {
                        dGV_FileList.Rows.Remove(row);
                    }
                    dGV_FileList.Refresh();
                }
            }
            else
            {
                MessageBox.Show("Please select the row(s) to delete.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void btn_reset_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Would you like to reset filelist?", "Reset Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                Init_dGV();
            }
        }
    }
}
