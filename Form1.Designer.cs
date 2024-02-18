
namespace ExcelToWord
{
    partial class iron66
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(iron66));
            this.processBtn = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.fileLabel = new System.Windows.Forms.Label();
            this.uploadBtn = new System.Windows.Forms.Button();
            this.fileDirLabel = new System.Windows.Forms.Label();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.SuspendLayout();
            // 
            // processBtn
            // 
            resources.ApplyResources(this.processBtn, "processBtn");
            this.processBtn.Name = "processBtn";
            this.processBtn.UseVisualStyleBackColor = true;
            this.processBtn.Click += new System.EventHandler(this.processBtn_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            this.openFileDialog1.FileOk += new System.ComponentModel.CancelEventHandler(this.openFileDialog1_FileOk);
            // 
            // fileLabel
            // 
            resources.ApplyResources(this.fileLabel, "fileLabel");
            this.fileLabel.Name = "fileLabel";
            // 
            // uploadBtn
            // 
            this.uploadBtn.AllowDrop = true;
            resources.ApplyResources(this.uploadBtn, "uploadBtn");
            this.uploadBtn.Name = "uploadBtn";
            this.uploadBtn.UseVisualStyleBackColor = true;
            this.uploadBtn.Click += new System.EventHandler(this.uploadBtn_Click);
            // 
            // fileDirLabel
            // 
            resources.ApplyResources(this.fileDirLabel, "fileDirLabel");
            this.fileDirLabel.Name = "fileDirLabel";
            this.fileDirLabel.Click += new System.EventHandler(this.fileDirLabel_Click);
            // 
            // progressBar1
            // 
            resources.ApplyResources(this.progressBar1, "progressBar1");
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Click += new System.EventHandler(this.progressBar1_Click);
            // 
            // iron66
            // 
            this.AllowDrop = true;
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.fileDirLabel);
            this.Controls.Add(this.uploadBtn);
            this.Controls.Add(this.fileLabel);
            this.Controls.Add(this.processBtn);
            this.Cursor = System.Windows.Forms.Cursors.Default;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Name = "iron66";
            this.Load += new System.EventHandler(this.iron66_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button processBtn;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Label fileLabel;
        private System.Windows.Forms.Button uploadBtn;
        private System.Windows.Forms.Label fileDirLabel;
        private System.Windows.Forms.ProgressBar progressBar1;
    }
}

