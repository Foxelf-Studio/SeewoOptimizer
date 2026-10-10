namespace SeewoOpt
{
    /// <summary>
    /// SettingsForm 的 Designer 部分。
    ///
    /// 【为什么这里是空的】
    /// 原实现把每个控件的 Location/Size 硬编码在 InitializeComponent 里，
    /// 那是从设计器拖出来的典型产物。本次改动要加入"最多 5 组关机时间点"，
    /// 每组含时间输入、每天开关与七个星期按钮——这类**行数可变**的布局
    /// 用绝对坐标手排会非常脆弱：任何一个控件改宽，后面全要跟着挪。
    ///
    /// 因此布局改为在 SettingsForm.cs 的 BuildLayout() 里用代码构建，
    /// 行高由 RuleRowHeight 常量统一控制。这里只保留最小骨架，
    /// 不删除文件是为了不让 .csproj 的 <Compile> 项与文件系统脱节。
    /// </summary>
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
            this.components = new System.ComponentModel.Container();
            this.SuspendLayout();
            //
            // SettingsForm
            //
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Name = "SettingsForm";
            this.ResumeLayout(false);
        }

        #endregion
    }
}
