using System;
using System.Drawing;
using System.Windows.Forms;

namespace SeewoOpt
{
    /// <summary>
    /// 「即将关机提醒」对话框：两个选项——「本次不关机」与「确认」。
    ///
    /// 【为什么不用 MessageBox】
    /// MessageBox 的按钮文字是系统内置的，中文 Windows 上固定渲染成
    /// "是(Y)" / "否(N)"，没有 API 可以改。用户明确要求按钮写
    /// 「本次不关机」和「确认」——这两个词承载的是完整语义
    /// （"是/否"要用户自己去猜哪个对应不关机），在教室一体机这种
    /// 无人值守场景下必须一眼看懂。所以只能用自绘对话框。
    ///
    /// 【为什么不用 Designer 文件】
    /// 这个对话框只有标题、一段文字、两个按钮，控件全是代码里三行就能摆好的，
    /// 再维护一个 .Designer.cs + .resx 只会增加改一处要同步三处的负担。
    /// 布局全部在构造函数里用固定坐标完成，尺寸随文本量固定，不涉及缩放。
    ///
    /// 【为什么是静态 Show 风格而不是构造即显示】
    /// 保留标准 Form 用法（构造 → ShowDialog(owner)），
    /// DialogResult 的语义与 MessageBox 保持一致：
    /// Yes = 用户点了「确认」，No = 用户点了「本次不关机」。
    /// 这样调用方可以从 MessageBox 平滑迁移过来，判定逻辑不用改。
    /// </summary>
    public class ShutdownWarningDialog : Form
    {
        /// <summary>
        /// 构造对话框。
        /// </summary>
        /// <param name="minutesAhead">还有多少分钟关机（提前量）</param>
        /// <param name="dueAt">预定关机的绝对时刻，用于在正文里写出具体时间</param>
        public ShutdownWarningDialog(int minutesAhead, DateTime dueAt)
        {
            // ---------- 窗体本身 ----------
            this.Text = "即将关机提醒";
            this.FormBorderStyle = FormBorderStyle.FixedDialog;   // 不可随意拉伸
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.ShowInTaskbar = true;                            // 一体机上要能被点到
            this.TopMost = true;                                  // 别被全屏课件压在下面
            this.BackColor = Color.White;
            this.Font = new Font("Microsoft YaHei UI", 9F);
            this.ClientSize = new Size(430, 190);

            // ---------- 警告图标 ----------
            // 用系统自带的信息图标，不额外嵌入资源。
            var icon = new PictureBox
            {
                Location = new Point(22, 26),
                Size = new Size(32, 32),
                SizeMode = PictureBoxSizeMode.StretchImage,
                Image = SystemIcons.Warning.ToBitmap()
            };
            this.Controls.Add(icon);

            // ---------- 正文 ----------
            var body = new Label
            {
                Location = new Point(68, 24),
                Size = new Size(340, 96),
                AutoSize = false,
                Font = new Font("Microsoft YaHei UI", 10F),
                ForeColor = Color.FromArgb(32, 32, 32),
                Text = string.Format(
                    "电脑将在 {0} 分钟后关机（{1:HH:mm}）。\r\n\r\n" +
                    "如果现在选择「本次不关机」，这一次不会关机，\r\n" +
                    "但下一个预定时间仍会照常执行。",
                    minutesAhead, dueAt)
            };
            this.Controls.Add(body);

            // ---------- 按钮 ----------
            // 尺寸参照系统按钮的常见比例（88x30），中文四字词足够容纳。
            const int btnW = 104;
            const int btnH = 32;
            const int gap = 14;
            int btnY = this.ClientSize.Height - btnH - 22;
            // 两个按钮整体靠右排布，与 Windows 习惯一致
            int confirmX = this.ClientSize.Width - 22 - btnW;
            int skipX = confirmX - gap - btnW;

            var btnSkip = new Button
            {
                Text = "本次不关机",
                Location = new Point(skipX, btnY),
                Size = new Size(btnW, btnH),
                DialogResult = DialogResult.No,
                // 【默认按钮放在"本次不关机"上】
                // 误按回车最坏的结果应该是"这次不关机"，而不是关机。
                // 这与原先 MessageBoxDefaultButton.Button2 的取向一致。
                TabIndex = 0
            };
            var btnConfirm = new Button
            {
                Text = "确认",
                Location = new Point(confirmX, btnY),
                Size = new Size(btnW, btnH),
                DialogResult = DialogResult.Yes,
                TabIndex = 1
            };
            this.Controls.Add(btnSkip);
            this.Controls.Add(btnConfirm);

            this.AcceptButton = btnSkip;   // 回车 = 本次不关机（安全侧）
            this.CancelButton = btnSkip;   // Esc 也走安全侧，避免手滑关机
        }
    }
}
