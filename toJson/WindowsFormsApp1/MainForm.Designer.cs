
namespace WindowsFormsApp1
{
    partial class MainForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
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
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.btn_load = new System.Windows.Forms.Button();
            this.ofd = new System.Windows.Forms.OpenFileDialog();
            this.bnt_convert = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btn_load
            // 
            this.btn_load.Location = new System.Drawing.Point(37, 32);
            this.btn_load.Name = "btn_load";
            this.btn_load.Size = new System.Drawing.Size(94, 29);
            this.btn_load.TabIndex = 0;
            this.btn_load.Text = "Select file";
            this.btn_load.UseVisualStyleBackColor = true;
            this.btn_load.Click += new System.EventHandler(this.btn_load_Click);
            // 
            // ofd
            // 
            this.ofd.FileName = "Select a file";
            this.ofd.Filter = "Hwp files (*.hwp)|*.hwp";
            this.ofd.Title = "Open hwp file";
            // 
            // bnt_convert
            // 
            this.bnt_convert.Location = new System.Drawing.Point(37, 106);
            this.bnt_convert.Name = "bnt_convert";
            this.bnt_convert.Size = new System.Drawing.Size(130, 29);
            this.bnt_convert.TabIndex = 1;
            this.bnt_convert.Text = "Convert File";
            this.bnt_convert.UseVisualStyleBackColor = true;
            this.bnt_convert.Click += new System.EventHandler(this.bnt_convert_Click);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.bnt_convert);
            this.Controls.Add(this.btn_load);
            this.Name = "MainForm";
            this.Text = "MainForm";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btn_load;
        private System.Windows.Forms.OpenFileDialog ofd;
        private System.Windows.Forms.Button bnt_convert;
    }
}