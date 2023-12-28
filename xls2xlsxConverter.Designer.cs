using xls2xlsxConverter.Properties;

namespace xls2xlsxConverter
{
    partial class xls2xlsxConverter
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(xls2xlsxConverter));
            dGV_FileList = new DataGridView();
            btn_convert = new Button();
            openFileDialog1 = new OpenFileDialog();
            btn_FileAdd = new Button();
            btn_FileDelete = new Button();
            btn_reset = new Button();
            btn_folder = new Button();
            txt_selectedPath = new TextBox();
            label_SaveDirectory = new Label();
            ((System.ComponentModel.ISupportInitialize)dGV_FileList).BeginInit();
            SuspendLayout();
            // 
            // dGV_FileList
            // 
            dGV_FileList.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dGV_FileList.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dGV_FileList.Location = new Point(14, 92);
            dGV_FileList.Name = "dGV_FileList";
            dGV_FileList.ReadOnly = true;
            dGV_FileList.RowHeadersWidth = 62;
            dGV_FileList.RowTemplate.Height = 33;
            dGV_FileList.Size = new Size(672, 447);
            dGV_FileList.TabIndex = 0;
            // 
            // btn_convert
            // 
            btn_convert.Image = Resources.xlsxicon;
            btn_convert.ImageAlign = ContentAlignment.TopCenter;
            btn_convert.Location = new Point(692, 442);
            btn_convert.Name = "btn_convert";
            btn_convert.Size = new Size(101, 97);
            btn_convert.TabIndex = 2;
            btn_convert.Text = "Convert";
            btn_convert.TextAlign = ContentAlignment.BottomCenter;
            btn_convert.UseVisualStyleBackColor = true;
            btn_convert.Click += btn_convert_Click;
            // 
            // openFileDialog1
            // 
            openFileDialog1.FileName = "openFileDialog1";
            // 
            // btn_FileAdd
            // 
            btn_FileAdd.Location = new Point(11, 49);
            btn_FileAdd.Name = "btn_FileAdd";
            btn_FileAdd.Size = new Size(112, 34);
            btn_FileAdd.TabIndex = 3;
            btn_FileAdd.Text = "File Add";
            btn_FileAdd.UseVisualStyleBackColor = true;
            btn_FileAdd.Click += btn_FileAdd_Click;
            // 
            // btn_FileDelete
            // 
            btn_FileDelete.Location = new Point(129, 49);
            btn_FileDelete.Name = "btn_FileDelete";
            btn_FileDelete.Size = new Size(112, 34);
            btn_FileDelete.TabIndex = 4;
            btn_FileDelete.Text = "File Delete";
            btn_FileDelete.UseVisualStyleBackColor = true;
            btn_FileDelete.Click += btn_FileDelete_Click;
            // 
            // btn_reset
            // 
            btn_reset.Location = new Point(247, 49);
            btn_reset.Name = "btn_reset";
            btn_reset.Size = new Size(112, 34);
            btn_reset.TabIndex = 5;
            btn_reset.Text = "Reset";
            btn_reset.UseVisualStyleBackColor = true;
            btn_reset.Click += btn_reset_Click;
            // 
            // btn_folder
            // 
            btn_folder.Location = new Point(650, 7);
            btn_folder.Name = "btn_folder";
            btn_folder.Size = new Size(139, 34);
            btn_folder.TabIndex = 6;
            btn_folder.Text = "Choose Folder";
            btn_folder.UseVisualStyleBackColor = true;
            btn_folder.Click += btn_folder_Click;
            // 
            // txt_selectedPath
            // 
            txt_selectedPath.Location = new Point(127, 10);
            txt_selectedPath.Name = "txt_selectedPath";
            txt_selectedPath.Size = new Size(518, 31);
            txt_selectedPath.TabIndex = 7;
            // 
            // label_SaveDirectory
            // 
            label_SaveDirectory.AutoSize = true;
            label_SaveDirectory.Location = new Point(14, 12);
            label_SaveDirectory.Name = "label_SaveDirectory";
            label_SaveDirectory.Size = new Size(107, 25);
            label_SaveDirectory.TabIndex = 8;
            label_SaveDirectory.Text = "Save Folder";
            // 
            // xls2xlsxConverter
            // 
            AutoScaleDimensions = new SizeF(10F, 25F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(798, 548);
            Controls.Add(label_SaveDirectory);
            Controls.Add(txt_selectedPath);
            Controls.Add(btn_folder);
            Controls.Add(btn_reset);
            Controls.Add(btn_FileDelete);
            Controls.Add(btn_FileAdd);
            Controls.Add(btn_convert);
            Controls.Add(dGV_FileList);
            Icon = (Icon)resources.GetObject("$this.Icon");
            Name = "xls2xlsxConverter";
            Text = "xls to xlsx Converter";
            ((System.ComponentModel.ISupportInitialize)dGV_FileList).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private DataGridView dGV_FileList;
        private Button btn_convert;
        private OpenFileDialog openFileDialog1;
        private Button btn_FileAdd;
        private Button btn_FileDelete;
        private Button btn_reset;
        private Button btn_folder;
        private TextBox txt_selectedPath;
        private Label label_SaveDirectory;
    }
}
