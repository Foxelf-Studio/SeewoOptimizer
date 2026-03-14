namespace TimeSyncTool
{
    partial class SettingsForm
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        private void InitializeComponent()
        {
            this.chkAutoStart = new System.Windows.Forms.CheckBox();
            this.chkSilentStart = new System.Windows.Forms.CheckBox();
            this.chkAutoVolume = new System.Windows.Forms.CheckBox();
            this.numVolumeLevel = new System.Windows.Forms.NumericUpDown();
            this.chkKillWps = new System.Windows.Forms.CheckBox();
            this.btnSave = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.numVolumeLevel)).BeginInit();
            this.SuspendLayout();

            // chkAutoStart
            this.chkAutoStart.AutoSize = true;
            this.chkAutoStart.Location = new System.Drawing.Point(12, 12);
            this.chkAutoStart.Size = new System.Drawing.Size(96, 16);
            this.chkAutoStart.Text = "开机自启动";
            this.chkAutoStart.UseVisualStyleBackColor = true;

            // chkSilentStart
            this.chkSilentStart.AutoSize = true;
            this.chkSilentStart.Location = new System.Drawing.Point(12, 34);
            this.chkSilentStart.Size = new System.Drawing.Size(84, 16);
            this.chkSilentStart.Text = "静默启动";
            this.chkSilentStart.UseVisualStyleBackColor = true;

            // chkAutoVolume
            this.chkAutoVolume.AutoSize = true;
            this.chkAutoVolume.Location = new System.Drawing.Point(12, 56);
            this.chkAutoVolume.Size = new System.Drawing.Size(120, 16);
            this.chkAutoVolume.Text = "自动执行音量调节";
            this.chkAutoVolume.UseVisualStyleBackColor = true;

            // numVolumeLevel
            this.numVolumeLevel.Location = new System.Drawing.Point(140, 54);
            this.numVolumeLevel.Size = new System.Drawing.Size(60, 21);
            this.numVolumeLevel.Maximum = 100;
            this.numVolumeLevel.Minimum = 0;
            this.numVolumeLevel.Value = 60;

            // chkKillWps
            this.chkKillWps.AutoSize = true;
            this.chkKillWps.Location = new System.Drawing.Point(12, 78);
            this.chkKillWps.Size = new System.Drawing.Size(162, 16);
            this.chkKillWps.Text = "自动执行杀死WPS进程任务";
            this.chkKillWps.UseVisualStyleBackColor = true;

            // btnSave
            this.btnSave.Location = new System.Drawing.Point(140, 100);
            this.btnSave.Size = new System.Drawing.Size(75, 23);
            this.btnSave.Text = "保存";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);

            // SettingsForm
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(220, 135);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.chkKillWps);
            this.Controls.Add(this.numVolumeLevel);
            this.Controls.Add(this.chkAutoVolume);
            this.Controls.Add(this.chkSilentStart);
            this.Controls.Add(this.chkAutoStart);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Name = "SettingsForm";
            this.Text = "系统优化设置";
            ((System.ComponentModel.ISupportInitialize)(this.numVolumeLevel)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion

        private System.Windows.Forms.CheckBox chkAutoStart;
        private System.Windows.Forms.CheckBox chkSilentStart;
        private System.Windows.Forms.CheckBox chkAutoVolume;
        private System.Windows.Forms.NumericUpDown numVolumeLevel;
        private System.Windows.Forms.CheckBox chkKillWps;
        private System.Windows.Forms.Button btnSave;
    }
}