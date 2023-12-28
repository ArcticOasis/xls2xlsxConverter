using System.Data;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
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
            dt.Columns.Add("InputFilePath", typeof(string));
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
            using (var OFD = new OpenFileDialog())
            {
                OFD.Filter = "엑셀 파일 (*.xls)|*.xls";
                DialogResult result = OFD.ShowDialog();

                if (result == DialogResult.OK) 
                {
                    MessageBox.Show(OFD.FileName, "Happy", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

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

        private void btn_convert_Click(object sender, EventArgs e)
        {
            if (!String.IsNullOrEmpty(txt_selectedPath.Text))
            {
                foreach (DataGridViewRow row in dGV_FileList.Rows)
                {
                    if (!row.IsNewRow)
                    {
                        string inputFilePath = row.Cells["InputFilePath"].Value.ToString(); // 열의 이름은 실제 설정된 열 이름에 따라 바꿔주세요
                        string outputFilePath = txt_selectedPath.Text + row.Cells["FileName"].Value.ToString() + ".xlsx"; // 원하는 방식으로 출력 파일 경로를 설정해주세요

                        ConvertXlsToXlsx(inputFilePath, outputFilePath);
                    }
                }
            }
            else
            {
                MessageBox.Show("Please check the save file path.", "File path is empty", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        // Logic Convert Xls => Xlsx
        private void ConvertXlsToXlsx(string inputFilePath, string outputFilePath)
        {
            try
            {
                using (var fs = new FileStream(inputFilePath, FileMode.Open, FileAccess.Read))
                {
                    IWorkbook workbook;
                    if (Path.GetExtension(inputFilePath).Equals(".xls"))
                    {
                        workbook = new HSSFWorkbook(fs); // HSSFWorkbook for xls
                    }
                    else
                    {
                        workbook = new XSSFWorkbook(fs); // XSSFWorkbook for xlsx
                    }

                    using (var output = new FileStream(outputFilePath, FileMode.Create, FileAccess.Write))
                    {
                        workbook.Write(output);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
