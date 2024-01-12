namespace LecturaExcel.View
{
    partial class Load_File
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
            this.Btn_Load_File = new MaterialSkin.Controls.MaterialRaisedButton();
            this.SuspendLayout();
            // 
            // Btn_Load_File
            // 
            this.Btn_Load_File.Depth = 0;
            this.Btn_Load_File.Location = new System.Drawing.Point(109, 165);
            this.Btn_Load_File.MouseState = MaterialSkin.MouseState.HOVER;
            this.Btn_Load_File.Name = "Btn_Load_File";
            this.Btn_Load_File.Primary = true;
            this.Btn_Load_File.Size = new System.Drawing.Size(154, 43);
            this.Btn_Load_File.TabIndex = 2;
            this.Btn_Load_File.Text = "Load File";
            this.Btn_Load_File.UseVisualStyleBackColor = true;
            this.Btn_Load_File.Click += new System.EventHandler(this.Btn_Load_File_Click);
            // 
            // Load_File
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.ClientSize = new System.Drawing.Size(397, 346);
            this.Controls.Add(this.Btn_Load_File);
            this.ForeColor = System.Drawing.SystemColors.ActiveCaption;
            this.Name = "Load_File";
            this.Sizable = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Load += new System.EventHandler(this.Load_File_Load);
            this.ResumeLayout(false);

        }

        #endregion
        private MaterialSkin.Controls.MaterialRaisedButton Btn_Load_File;
    }
}