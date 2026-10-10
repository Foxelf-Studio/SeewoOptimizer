using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
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
    /// 【为什么控件全部自绘，连按钮都不用原生 Button】
    /// 第一版用了原生 Button + Microsoft YaHei UI，实机截图被用户指出
    /// "布局和字体有点山寨 Windows"。根因是**两套视觉语言混搭**：
    /// 自绘的窗体背景/正文是雅黑，而原生 Button 用的是系统默认字体
    /// （中文 Windows 上是宋体 9pt），标题栏又走系统绘制——三者同屏，
    /// 一眼就看得出不是一个时代的东西。
    ///
    /// 更根本的问题是：**原生 Button 在 Win7 上不支持自定义配色**。
    /// 设了 BackColor 会走"视觉样式已禁用"的丑化路径（灰色方块），
    /// 不设又跟自绘背景不搭。所以这里做「全自绘」：
    /// 窗体背景、图标、正文、按钮全部自己画，
    /// 控件树里只有 Label 与自绘 Button 两种，字体统一为 Microsoft YaHei。
    ///
    /// 【为什么不用 Designer 文件】
    /// 控件全是代码里能摆好的，再维护 .Designer.cs + .resx 只会
    /// 增加"改一处要同步三处"的负担。布局在构造函数里一次算完。
    ///
    /// 【为什么保留标准 Form 用法（构造 → ShowDialog）】
    /// DialogResult 的语义与 MessageBox 保持一致：
    /// Yes = 用户点了「确认」，No = 用户点了「本次不关机」。
    /// 这样调用方从 MessageBox 迁移过来时判定逻辑一行都不用改。
    /// </summary>
    public class ShutdownWarningDialog : Form
    {
        // ============================================================
        // 视觉常量
        //
        // 【为什么集中放在这里】第一版把字体、颜色、尺寸散落在各处
        // （字体写了 4 处、颜色硬编码 6 处），改一处配色要翻遍整个文件，
        // 极易漏改导致不一致。集中定义后"改主题"是改这几个常量，
        // 而不是在几百行里找散落的魔数。
        // ============================================================

        /// <summary>
        /// 统一字体族。
        ///
        /// 【为什么是 Microsoft YaHei 而不是 "Microsoft YaHei UI"】
        /// 后者是 Win8.1 起才随系统附带的 UI 变体，在 Win7 上不存在时
        /// 会静默回退到宋体，字形与其他控件不一致。本项目其余界面
        /// （TimeSyncForm / SettingsForm）用的都是 Microsoft YaHei，
        /// 这里保持一致。教室一体机多为 Win7/Win10 混合，
        /// Microsoft YaHei 两边都有，是最稳的选择。
        /// </summary>
        private const string FontFamily = "Microsoft YaHei";

        /// <summary>标准正文：9.5pt 是中文界面下可读性与紧凑度的平衡点</summary>
        private const float BodySize = 9.5f;

        /// <summary>标题/强调文字</summary>
        private const float TitleSize = 10.5f;

        // 配色：取 Windows 对话框的灰阶体系，而不是纯白/纯黑。
        // 纯白背景 + 纯黑文字在 Aero 主题下对比过强，反而显廉价。
        private static readonly Color FormBack = Color.FromArgb(240, 240, 240);   // 系统对话框底色
        private static readonly Color TextPrimary = Color.FromArgb(26, 26, 26);   // 正文（近黑，非纯黑）
        private static readonly Color TextSecondary = Color.FromArgb(96, 96, 96); // 次要说明
        private static readonly Color AccentBlue = Color.FromArgb(30, 90, 160);   // 与主窗体一致的主题蓝（用于按钮悬停态）

        /// <summary>
        /// 全局字体。构造前静态初始化一次。
        ///
        /// 【为什么用静态只读而不是每处 new Font()】
        /// GDI 字体对象是有限句柄资源；对话框可能被反复弹出（每次提醒一个），
        /// 每处 new 而不 Dispose 会累积泄漏。静态共享一份，
        /// 随进程存活，不涉及释放问题。
        /// </summary>
        private static readonly Font BodyFont = new Font(FontFamily, BodySize);
        private static readonly Font TitleFont = new Font(FontFamily, TitleSize, FontStyle.Bold);
        private static readonly Font ButtonFont = new Font(FontFamily, BodySize);

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
            this.BackColor = FormBack;
            // 字体设在窗体上，未显式指定字体的子控件会继承。
            // 这是"字体统一"的兜底——漏设一处也不会掉回宋体。
            this.Font = BodyFont;
            // 关掉系统默认的图标缩放（否则 DPI 放大时系统图标会二次插值变糊）
            this.AutoScaleMode = AutoScaleMode.None;
            this.ClientSize = new Size(430, 196);

            // ---------- 警告图标（自绘，不用 SystemIcons） ----------
            // SystemIcons.Warning 是 32×32 的老式位图，拉伸到 48 会糊。
            // 这里按 40×40 自己画一个干净的三色警示三角。
            //
            // 【为什么是 40 而不是 44】44 时三角在浅灰底上显得笨重，
            // 40 与右侧两行文字的视觉高度更接近，整体更平衡。
            var icon = new PictureBox
            {
                Location = new Point(24, 26),
                Size = new Size(40, 40),
                SizeMode = PictureBoxSizeMode.Zoom,
                Image = BuildWarningBitmap(40)
            };
            this.Controls.Add(icon);

            // ---------- 标题 ----------
            var title = new Label
            {
                Text = "电脑即将关机",
                Location = new Point(80, 26),
                Size = new Size(330, 22),
                Font = TitleFont,
                ForeColor = TextPrimary,
                BackColor = Color.Transparent
            };
            this.Controls.Add(title);

            // ---------- 正文第一段（含具体时刻） ----------
            var line1 = new Label
            {
                Location = new Point(80, 52),
                Size = new Size(334, 20),
                Font = BodyFont,
                ForeColor = TextPrimary,
                BackColor = Color.Transparent,
                Text = string.Format("电脑将在 {0} 分钟后关机（{1:HH:mm}）。",
                                     minutesAhead, dueAt)
            };
            this.Controls.Add(line1);

            // ---------- 正文第二段（说明"本次不关机"的含义） ----------
            //
            // 【为什么拆成独立 Label 而不是一个多行 Label 里塞 \r\n】
            // 第一版用硬编码换行符折行，换字体或换 DPI 后折行位置就错了，
            // 表现为"参差不齐"。拆成两个 Label 后，一段文字只表达一个意思，
            // 换行由各自的视觉边界决定，不依赖字符宽度。
            //
            // 【为什么要手动折行而不是靠 MaximumSize 自动折】
            // 自动折行会在任意字符处断开（实测在"下"字前断成"…但下 / 一个…"），
            // 中文里这种断法读起来别扭。这里按语义手动断在"这一次不会关机，"之后，
            // 两行长度接近，视觉更整齐。宽度用 MaximumSize 约束兜底，
            // 万一将来文案变长仍有自动折行保护，不会溢出。
            var line2 = new Label
            {
                Location = new Point(80, 78),
                Font = BodyFont,
                ForeColor = TextSecondary,
                BackColor = Color.Transparent,
                AutoSize = true,
                MaximumSize = new Size(336, 0),
                Text = "如果现在选择「本次不关机」，这一次不会关机，" + Environment.NewLine
                     + "但下一个预定时间仍会照常执行。"
            };
            this.Controls.Add(line2);

            // ---------- 按钮 ----------
            // 自绘按钮：原生 Button 在 Win7 上无法同时做到"自定义配色"和"正常外观"。
            const int btnW = 104;
            const int btnH = 30;
            const int gap = 10;
            int btnY = this.ClientSize.Height - btnH - 18;
            int confirmX = this.ClientSize.Width - 24 - btnW;
            int skipX = confirmX - gap - btnW;

            var btnSkip = new FlatButton
            {
                Text = "本次不关机",
                Location = new Point(skipX, btnY),
                Size = new Size(btnW, btnH),
                DialogResult = DialogResult.No,
                Font = ButtonFont,
                // 【默认按钮放在"本次不关机"上】
                // 误按回车最坏的结果应该是"这次不关机"，而不是关机。
                // 与原先 MessageBoxDefaultButton.Button2 的取向一致。
                TabIndex = 0
            };
            var btnConfirm = new FlatButton
            {
                Text = "确认",
                Location = new Point(confirmX, btnY),
                Size = new Size(btnW, btnH),
                DialogResult = DialogResult.Yes,
                Font = ButtonFont,
                TabIndex = 1
            };
            this.Controls.Add(btnSkip);
            this.Controls.Add(btnConfirm);

            this.AcceptButton = btnSkip;   // 回车 = 本次不关机（安全侧）
            this.CancelButton = btnSkip;   // Esc 也走安全侧，避免手滑关机
        }

        /// <summary>
        /// 生成一个干净的警示三角图标。
        ///
        /// 【为什么自己画而不是用 SystemIcons.Warning】
        /// 系统图标是固定 32×32 的老式位图，放到 40px 会明显发虚；
        /// 且它是黄色实心三角配黑色感叹号，在浅灰背景上偏"廉价提示音"的观感。
        /// 自绘可以做到：圆角顶点、抗锯齿、与整体灰阶协调的琥珀黄。
        ///
        /// 【比例上的几个取值理由】
        ///   · 三角整体缩到 0.86 倍并居中，四周留白 —— 贴边的图标显拥挤
        ///   · 描边取 0.045 倍宽 —— 0.055 时边框过粗，三角看起来像实心色块
        ///   · 感叹号整体上移 —— 三角底边是斜的，视觉重心偏下，
        ///     若按几何居中会让感叹号看起来"往下掉"
        /// </summary>
        private static Bitmap BuildWarningBitmap(int size)
        {
            var bmp = new Bitmap(size, size);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;
                g.Clear(Color.Transparent);

                float unit = size / 100f;              // 以 100 为基准单位
                float cx = size / 2f;

                // 三角顶点：整体缩到 0.86 并居中
                float top = 8f * unit;
                float bottom = 92f * unit;
                float half = 42f * unit;
                var p1 = new PointF(cx, top);
                var p2 = new PointF(cx + half, bottom);
                var p3 = new PointF(cx - half, bottom);

                using (var tri = new GraphicsPath())
                {
                    tri.AddPolygon(new[] { p1, p2, p3 });
                    using (var fill = new SolidBrush(Color.FromArgb(255, 201, 71)))
                        g.FillPath(fill, tri);
                    using (var pen = new Pen(Color.FromArgb(212, 150, 25), 4.5f * unit))
                    {
                        pen.LineJoin = LineJoin.Round;
                        g.DrawPath(pen, tri);
                    }
                }

                // 感叹号：竖线 + 圆点，整体略上移
                using (var fg = new SolidBrush(Color.FromArgb(78, 56, 0)))
                using (var stem = new Pen(Color.FromArgb(78, 56, 0), 8.5f * unit))
                {
                    stem.StartCap = LineCap.Round;
                    stem.EndCap = LineCap.Round;
                    g.DrawLine(stem, cx, 42f * unit, cx, 62f * unit);
                    float r = 4.6f * unit;
                    g.FillEllipse(fg, cx - r, 73f * unit - r, r * 2, r * 2);
                }
            }
            return bmp;
        }

        /// <summary>
        /// 扁平风格按钮：圆角描边 + 悬停高亮，外观贴近现代 Windows 对话框，
        /// 且完全由代码绘制，不依赖系统视觉样式。
        ///
        /// 【为什么不用原生 Button + FlatStyle.Flat】
        /// FlatStyle.Flat 在 Win7 上会走 "视觉样式已禁用" 的绘制分支，
        /// 鼠标悬停没有任何反馈，看起来像被禁用状态。
        /// 自绘可以精确控制三态（常态/悬停/按下）的描边与底色。
        ///
        /// 【为什么重写 OnPaint 而不是用区域裁剪】
        /// 小尺寸圆角用 Region 裁剪会有明显锯齿（Region 是硬边）；
        /// 用 GraphicsPath + AntiAlias 描边能得到平滑边缘。
        /// </summary>
        private class FlatButton : Button
        {
            private bool _hover;
            private bool _pressed;

            public FlatButton()
            {
                SetStyle(ControlStyles.AllPaintingInWmPaint
                       | ControlStyles.UserPaint
                       | ControlStyles.OptimizedDoubleBuffer
                       | ControlStyles.ResizeRedraw, true);
                FlatStyle = FlatStyle.Flat;
                FlatAppearance.BorderSize = 0;
                UseVisualStyleBackColor = false;
                Cursor = Cursors.Hand;
                BackColor = FormBack;   // 供透明底绘制使用
            }

            protected override void OnMouseEnter(EventArgs e)
            {
                _hover = true; Invalidate(); base.OnMouseEnter(e);
            }

            protected override void OnMouseLeave(EventArgs e)
            {
                _hover = false; _pressed = false; Invalidate(); base.OnMouseLeave(e);
            }

            protected override void OnMouseDown(MouseEventArgs e)
            {
                _pressed = true; Invalidate(); base.OnMouseDown(e);
            }

            protected override void OnMouseUp(MouseEventArgs e)
            {
                _pressed = false; Invalidate(); base.OnMouseUp(e);
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                Graphics g = e.Graphics;
                g.SmoothingMode = SmoothingMode.AntiAlias;

                // 与父窗体同色打底，避免出现"按钮自身背景块"
                using (var back = new SolidBrush(FormBack))
                    g.FillRectangle(back, ClientRectangle);

                // 三态配色
                Color fill, border, text;
                if (_pressed)
                {
                    fill = Color.FromArgb(214, 228, 244);
                    border = Color.FromArgb(90, 130, 180);
                    text = Color.FromArgb(20, 60, 110);
                }
                else if (_hover)
                {
                    fill = Color.FromArgb(232, 242, 253);
                    border = Color.FromArgb(120, 165, 215);
                    text = AccentBlue;
                }
                else
                {
                    fill = Color.FromArgb(253, 253, 253);
                    border = Color.FromArgb(172, 172, 172);
                    text = TextPrimary;
                }

                var rect = new Rectangle(0, 0, Width - 1, Height - 1);
                const int radius = 4;
                using (var path = RoundedRect(rect, radius))
                {
                    using (var b = new SolidBrush(fill)) g.FillPath(b, path);
                    using (var p = new Pen(border)) g.DrawPath(p, path);
                }

                TextRenderer.DrawText(
                    g, Text, Font, rect, text,
                    TextFormatFlags.HorizontalCenter
                  | TextFormatFlags.VerticalCenter
                  | TextFormatFlags.NoPadding);
            }

            /// <summary>画圆角矩形路径</summary>
            private static GraphicsPath RoundedRect(Rectangle r, int radius)
            {
                int d = radius * 2;
                var path = new GraphicsPath();
                path.AddArc(r.X, r.Y, d, d, 180, 90);
                path.AddArc(r.Right - d, r.Y, d, d, 270, 90);
                path.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
                path.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
                path.CloseFigure();
                return path;
            }
        }
    }
}
