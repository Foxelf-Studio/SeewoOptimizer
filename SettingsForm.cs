using System;
using System.Windows.Forms;

namespace TimeSyncTool
{
    public partial class SettingsForm : Form
    {
        private TimeSyncForm parentForm;

        public SettingsForm(TimeSyncForm parent)
        {
            InitializeComponent();

            parentForm = parent;

            // 加载当前设置
            chkAutoStart.Checked = parentForm.AutoStart;
            chkSilentStart.Checked = parentForm.SilentStart;
            chkAutoVolume.Checked = parentForm.AutoVolume;
            numVolumeLevel.Value = parentForm.VolumeLevel;
            chkKillWps.Checked = parentForm.KillWps;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            // 保存设置
            parentForm.AutoStart = chkAutoStart.Checked;
            parentForm.SilentStart = chkSilentStart.Checked;
            parentForm.AutoVolume = chkAutoVolume.Checked;
            parentForm.VolumeLevel = (int)numVolumeLevel.Value;
            parentForm.KillWps = chkKillWps.Checked;
            parentForm.SaveSettings();

            this.Close();
        }
    }
}