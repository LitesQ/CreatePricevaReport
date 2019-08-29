namespace Exel
{
    partial class Form1
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
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.Action3 = new System.Windows.Forms.Button();
            this.Status = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // Action3
            // 
            this.Action3.BackColor = System.Drawing.Color.DarkRed;
            this.Action3.Font = new System.Drawing.Font("Microsoft Sans Serif", 48F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.Action3.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.Action3.Location = new System.Drawing.Point(12, 12);
            this.Action3.Name = "Action3";
            this.Action3.Size = new System.Drawing.Size(664, 316);
            this.Action3.TabIndex = 4;
            this.Action3.Text = "БКК";
            this.Action3.UseVisualStyleBackColor = false;
            this.Action3.Click += new System.EventHandler(this.Action3_Click);
            // 
            // Status
            // 
            this.Status.AutoSize = true;
            this.Status.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.Status.Location = new System.Drawing.Point(12, 340);
            this.Status.Name = "Status";
            this.Status.Size = new System.Drawing.Size(0, 16);
            this.Status.TabIndex = 5;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.ClientSize = new System.Drawing.Size(688, 388);
            this.Controls.Add(this.Status);
            this.Controls.Add(this.Action3);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.ShowIcon = false;
            this.Text = "Отчет по ценам";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button Action3;
        private System.Windows.Forms.Label Status;
    }
}

