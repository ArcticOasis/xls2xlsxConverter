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
        }

        private void btn_folder_Click(object sender, EventArgs e)
        {
            using (var v_FBD = new FolderBrowserDialog())
            {
                DialogResult result = v_FBD.ShowDialog();

                if (result == DialogResult.OK)
                {
                    txt_selectedPath.Text = v_FBD.SelectedPath + @"\";
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
                    dGV_FileList.Rows.Add(dGV_FileList.Rows.Count + 1, OFD.SafeFileName, "STANDBY", OFD.FileName);
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
                dGV_FileList.Rows.Clear();
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
                        string inputFilePath = row.Cells["InputFilePath"].Value.ToString(); 
                        string outputFilePath = txt_selectedPath.Text + row.Cells["FileName"].Value.ToString().Replace(".xls", "") + ".xlsx";

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
                using (FileStream fs = new FileStream(inputFilePath, FileMode.Open, FileAccess.Read))
                {
                    IWorkbook xlsxWorkbook = new XSSFWorkbook();

                    IWorkbook xlsWorkbook;

                    if (Path.GetExtension(inputFilePath).Equals(".xls"))
                    {
                        xlsWorkbook = new HSSFWorkbook(fs);

                        for (int i = 0; i < xlsWorkbook.NumberOfSheets; i++)
                        {
                            ISheet xlsSheet = xlsWorkbook.GetSheetAt(i);
                            ISheet xlsxSheet = xlsxWorkbook.CreateSheet(xlsSheet.SheetName);

                            for (int j = 0; j <= xlsSheet.LastRowNum; j++)
                            {
                                IRow xlsRow = xlsSheet.GetRow(j);
                                IRow xlsxRow = xlsxSheet.CreateRow(j);

                                if (xlsRow != null)
                                {
                                    for (int k = 0; k < xlsRow.LastCellNum; k++)
                                    {
                                        ICell xlsCell = xlsRow.GetCell(k);
                                        ICell xlsxCell = xlsxRow.CreateCell(k);

                                        if (xlsCell != null)
                                        {
                                            object cellValue = GetCellValue(xlsCell);

                                            if (cellValue != null)
                                            {
                                                switch (xlsCell.CellType)
                                                {
                                                    case CellType.Numeric:
                                                        if (DateUtil.IsCellDateFormatted(xlsCell))
                                                        {
                                                            xlsxCell.SetCellValue(xlsCell.DateCellValue);
                                                            xlsxCell.SetCellType(CellType.Formula);
                                                            xlsxCell.CellFormula = xlsCell.CellFormula;
                                                        }
                                                        else
                                                        {
                                                            xlsxCell.SetCellValue((double)cellValue);
                                                        }
                                                        break;

                                                    case CellType.String:
                                                        xlsxCell.SetCellValue((string)cellValue);
                                                        break;

                                                    case CellType.Boolean:
                                                        xlsxCell.SetCellValue((bool)cellValue);
                                                        break;

                                                    case CellType.Formula:
                                                        xlsxCell.SetCellFormula(xlsCell.CellFormula);
                                                        break;

                                                    case CellType.Blank:
                                                        break;

                                                    case CellType.Error:
                                                        xlsxCell.SetCellValue("Error Cell");
                                                        break;
                                                    default:
                                                        xlsxCell.SetCellValue("Data replication for this cell failed.");
                                                        break;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("This file is not xls", "This file is not xls", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }

                    MessageBox.Show(outputFilePath, "File path is empty", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                    using (FileStream output = new FileStream(outputFilePath, FileMode.Create, FileAccess.Write))
                    {
                        xlsxWorkbook.Write(output);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        // Cell Value Type Export
        static object GetCellValue(ICell cell)
        {
            switch (cell.CellType)
            {
                case CellType.Numeric:
                    return cell.NumericCellValue;
                case CellType.String:
                    return cell.StringCellValue;
                case CellType.Boolean:
                    return cell.BooleanCellValue;
                case CellType.Formula:
                    return cell.CellFormula;
                case CellType.Blank:
                case CellType.Error:
                default:
                    return null;
            }
        }

        static void CopyCellStyle(ICellStyle oldStyle, ICellStyle newStyle)
        {
            newStyle.Alignment = oldStyle.Alignment;
            newStyle.VerticalAlignment = oldStyle.VerticalAlignment;
            newStyle.WrapText = oldStyle.WrapText;
            newStyle.Indention = oldStyle.Indention;
            newStyle.Rotation = oldStyle.Rotation;
            newStyle.BorderBottom = oldStyle.BorderBottom;
            newStyle.BorderLeft = oldStyle.BorderLeft;
            newStyle.BorderRight = oldStyle.BorderRight;
            newStyle.BorderTop = oldStyle.BorderTop;
            newStyle.BottomBorderColor = oldStyle.BottomBorderColor;
            newStyle.LeftBorderColor = oldStyle.LeftBorderColor;
            newStyle.RightBorderColor = oldStyle.RightBorderColor;
            newStyle.TopBorderColor = oldStyle.TopBorderColor;
            newStyle.FillForegroundColor = oldStyle.FillForegroundColor;
            newStyle.FillPattern = oldStyle.FillPattern;
            // 여러 다른 스타일 속성 복사
        }
    }
}
