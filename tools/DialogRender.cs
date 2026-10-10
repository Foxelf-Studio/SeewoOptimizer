using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Windows.Forms;

/// <summary>
/// 把 ShutdownWarningDialog 渲染成 PNG，用于在没有真机 GUI 的情况下
/// 核对布局与字体。
///
/// 【为什么必须这么做，而不是"看起来没问题"】
/// 沙箱无法启动常驻 GUI 进程（进程会被外部清理），
/// 但"样式改好了"这种结论不能靠想象——必须拿到真实像素。
/// DrawToBitmap 走的是窗体的真实绘制路径（OnPaint / 子控件绘制），
/// 得到的就是真机上会显示的内容（除标题栏由系统绘制外）。
///
/// 【为什么同时渲染"悬停态"】
/// 自绘按钮的核心价值在于三态反馈。常态好看不代表悬停态不塌，
/// 这里通过反射触发 _hover 字段后重绘，把悬停态也落盘。
/// </summary>
internal static class DialogRender
{
    [STAThread]
    private static int Main(string[] args)
    {
        string outDir = args.Length > 0
            ? args[0]
            : Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "render");

        Directory.CreateDirectory(outDir);

        try
        {
            // 高 DPI 感知，避免渲染时被系统 DPI 虚拟化缩放导致模糊
            Application.SetCompatibleTextRenderingDefault(false);
            Application.EnableVisualStyles();

            using (var dlg = new SeewoOpt.ShutdownWarningDialog(
                       5, new DateTime(2026, 10, 11, 0, 25, 0)))
            {
                // 【为什么必须真的 Show 一次】DrawToBitmap 只重绘"已创建句柄
                // 且已完成布局"的控件。仅调 CreateControl() 时，FixedDialog
                // 的子控件尚未走完布局流程，绘制结果是空白客户区（实测如此）。
                // 用 Show 后立刻 Hide，让布局真正发生，再 DrawToBitmap 才有内容。
                dlg.Show();
                Application.DoEvents();
                dlg.Refresh();
                Application.DoEvents();

                // ---- 常态 ----
                string normalPath = Path.Combine(outDir, "dialog-normal.png");
                SaveBitmap(dlg, normalPath);
                Console.WriteLine("[render] " + normalPath);

                // ---- 悬停态：把两个按钮的 _hover 置 true 再重绘 ----
                SetHover(dlg, true);
                Application.DoEvents();
                string hoverPath = Path.Combine(outDir, "dialog-hover.png");
                SaveBitmap(dlg, hoverPath);
                Console.WriteLine("[render] " + hoverPath);

                dlg.Hide();
            }

            Console.WriteLine("OK");
            return 0;
        }
        catch (Exception ex)
        {
            Console.WriteLine("渲染失败：" + ex);
            return 1;
        }
    }

    /// <summary>用反射把窗体里所有 FlatButton 的 _hover 字段设为指定值</summary>
    private static void SetHover(Control root, bool value)
    {
        foreach (Control c in root.Controls)
        {
            if (c.GetType().Name == "FlatButton")
            {
                var f = c.GetType().GetField("_hover",
                    System.Reflection.BindingFlags.Instance
                  | System.Reflection.BindingFlags.NonPublic);
                if (f != null) f.SetValue(c, value);
                c.Invalidate();
                c.Update();
            }
            if (c.HasChildren) SetHover(c, value);
        }
        root.Refresh();
    }

    private static void SaveBitmap(Form form, string path)
    {
        // 用窗体自身尺寸（含边框）绘制，尽量贴近真机截图
        int w = form.Width;
        int h = form.Height;

        using (var bmp = new Bitmap(w, h))
        {
            // 先铺一层中性背景，模拟桌面，便于看清窗体边界
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.Clear(Color.FromArgb(88, 130, 180));
            }

            form.DrawToBitmap(bmp, new Rectangle(0, 0, w, h));

            using (var fs = new FileStream(path, FileMode.Create, FileAccess.Write))
                bmp.Save(fs, ImageFormat.Png);
        }
    }
}
