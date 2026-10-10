using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using SeewoOpt.Services;

namespace SeewoOpt
{
    /// <summary>
    /// 设置窗体。
    ///
    /// 【为什么要重写而不是在原 Designer 上加控件】
    /// 原来的设置窗体是"一堆绝对坐标的复选框 + 一个保存按钮"，
    /// 高 135px 就够。这次要加的东西（自动执行开关 + 最多 5 组关机时间点，
    /// 每组含时间输入、每天开关、七个星期按钮）用绝对坐标手排会非常脆弱：
    /// 任何一个控件改宽一点，后面全要跟着挪。
    ///
    /// 因此改为**用代码动态构建 + 分区**：每个关机时间点是一行自定义控件，
    /// 行高统一由常量控制，增删行只需重建列表，不必动其它控件的位置。
    /// 这里不用 Designer 文件，把布局全部写在 <see cref="BuildLayout"/> 里，
    /// 意图比 Designer 生成的代码清楚得多。
    /// </summary>
    public partial class SettingsForm : Form
    {
        private TimeSyncForm parentForm;

        // --- 上半部分：常规选项 ---
        private CheckBox chkAutoStart;
        private CheckBox chkSilentStart;
        private CheckBox chkAutoStartTask;
        private CheckBox chkAutoVolume;
        private NumericUpDown numVolumeLevel;
        private CheckBox chkKillWps;

        // --- 下半部分：关机时间点 ---
        private Panel rulesPanel;
        private Button btnAddRule;
        private Button btnRemoveRule;
        private Button btnSave;
        private Label emptyHint;

        /// <summary>关机规则编辑行（每行一个时间点）</summary>
        private List<RuleRow> ruleRows = new List<RuleRow>();

        /// <summary>一行的高度，增删行时按此重排</summary>
        private const int RuleRowHeight = 30;

        public SettingsForm(TimeSyncForm parent)
        {
            InitializeComponent();

            parentForm = parent;

            BuildLayout();
            LoadFromParent();
        }

        // =================================================================
        // 布局
        // =================================================================

        private void BuildLayout()
        {
            this.Text = "系统优化设置";
            this.ClientSize = new Size(560, 440);
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterParent;
            this.Font = new Font("Microsoft YaHei", 9);

            // ---------------- 常规选项区 ----------------
            GroupBox optionGroup = new GroupBox
            {
                Text = "常规",
                Location = new Point(12, 10),
                Size = new Size(536, 150)
            };

            chkAutoStart = NewCheckBox("开机自启动", new Point(16, 26));
            chkSilentStart = NewCheckBox("静默启动（启动时不显示窗口，直接进托盘）", new Point(16, 50));

            // 需求 1 的核心开关
            chkAutoStartTask = NewCheckBox("程序启动时自动开始执行既定任务", new Point(16, 74));

            chkAutoVolume = NewCheckBox("自动执行音量调节", new Point(16, 98));

            numVolumeLevel = new NumericUpDown
            {
                Location = new Point(230, 96),
                Size = new Size(60, 23),
                Minimum = 0,
                Maximum = 100,
                Value = 60
            };

            chkKillWps = NewCheckBox("自动执行结束 WPS 进程任务", new Point(16, 122));

            optionGroup.Controls.Add(chkAutoStart);
            optionGroup.Controls.Add(chkSilentStart);
            optionGroup.Controls.Add(chkAutoStartTask);
            optionGroup.Controls.Add(chkAutoVolume);
            optionGroup.Controls.Add(numVolumeLevel);
            optionGroup.Controls.Add(chkKillWps);
            this.Controls.Add(optionGroup);

            // ---------------- 关机时间点区 ----------------
            GroupBox shutdownGroup = new GroupBox
            {
                Text = "定时关机（最多 5 个时间点，执行前 5 分钟提醒）",
                Location = new Point(12, 168),
                Size = new Size(536, 224)
            };

            rulesPanel = new Panel
            {
                Location = new Point(12, 22),
                Size = new Size(512, RuleRowHeight * AppSettings.MaxShutdownRules),
                AutoScroll = true,
                BackColor = Color.White
            };

            emptyHint = new Label
            {
                Text = "尚未设置关机时间点。点击下方「添加时间点」开始设置。",
                Location = new Point(12, 26),
                Size = new Size(500, 24),
                ForeColor = Color.Gray
            };
            rulesPanel.Controls.Add(emptyHint);

            btnAddRule = new Button
            {
                Text = "添加时间点",
                Location = new Point(12, 22 + RuleRowHeight * AppSettings.MaxShutdownRules + 8),
                Size = new Size(110, 28)
            };
            btnAddRule.Click += (s, e) => AddRuleRow(7, 0, 0, true);

            btnRemoveRule = new Button
            {
                Text = "删除末行",
                Location = new Point(130, 22 + RuleRowHeight * AppSettings.MaxShutdownRules + 8),
                Size = new Size(90, 28)
            };
            btnRemoveRule.Click += (s, e) => RemoveLastRuleRow();

            shutdownGroup.Controls.Add(rulesPanel);
            shutdownGroup.Controls.Add(btnAddRule);
            shutdownGroup.Controls.Add(btnRemoveRule);
            this.Controls.Add(shutdownGroup);

            // ---------------- 保存按钮 ----------------
            btnSave = new Button
            {
                Text = "保存",
                Location = new Point(455, 400),
                Size = new Size(90, 30),
                DialogResult = DialogResult.None
            };
            btnSave.Click += btnSave_Click;
            this.Controls.Add(btnSave);

            Button btnCancel = new Button
            {
                Text = "取消",
                Location = new Point(358, 400),
                Size = new Size(90, 30)
            };
            btnCancel.Click += (s, e) => this.Close();
            this.Controls.Add(btnCancel);
        }

        private CheckBox NewCheckBox(string text, Point location)
        {
            return new CheckBox
            {
                Text = text,
                Location = location,
                AutoSize = true,
                UseVisualStyleBackColor = true
            };
        }

        // =================================================================
        // 关机时间点行
        // =================================================================

        /// <summary>
        /// 一个关机时间点的编辑行。
        ///
        /// 【为什么用"每天开关 + 星期多选"，而不是只放七个按钮】
        /// 用户最常用的配置就是"每天"。若只有七个按钮，想设"每天"
        /// 就得点七下，还得检查有没有漏掉某天。加一个"每天"开关后，
        /// 常态配置一次点击完成，只有真需要挑日子时才展开到七天。
        /// 勾上"每天"时把七个按钮全部置为选中并禁用——让界面直接
        /// 呈现出"就是全选了"的事实，而不是让它变成一个隐式状态。
        /// </summary>
        private class RuleRow : Panel
        {
            public NumericUpDown HourBox;
            public NumericUpDown MinuteBox;
            public CheckBox EveryDayCheck;
            public CheckBox[] DayChecks = new CheckBox[7];

            private static readonly string[] DayTexts = { "一", "二", "三", "四", "五", "六", "日" };

            public RuleRow(int index, ShutdownTimeRule rule)
            {
                this.Location = new Point(4, index * RuleRowHeight);
                this.Size = new Size(500, RuleRowHeight - 2);
                this.BackColor = index % 2 == 0 ? Color.FromArgb(248, 250, 253) : Color.White;

                Label idx = new Label
                {
                    Text = (index + 1) + ".",
                    Location = new Point(4, 6),
                    Size = new Size(22, 18),
                    ForeColor = Color.Gray
                };
                this.Controls.Add(idx);

                HourBox = new NumericUpDown
                {
                    Location = new Point(28, 3),
                    Size = new Size(46, 23),
                    Minimum = 0,
                    Maximum = 23,
                    Value = rule != null ? rule.Hour : 0
                };
                this.Controls.Add(HourBox);

                this.Controls.Add(new Label
                {
                    Text = "时",
                    Location = new Point(76, 6),
                    Size = new Size(18, 18)
                });

                MinuteBox = new NumericUpDown
                {
                    Location = new Point(94, 3),
                    Size = new Size(46, 23),
                    Minimum = 0,
                    Maximum = 59,
                    Value = rule != null ? rule.Minute : 0
                };
                this.Controls.Add(MinuteBox);

                this.Controls.Add(new Label
                {
                    Text = "分",
                    Location = new Point(142, 6),
                    Size = new Size(18, 18)
                });

                EveryDayCheck = new CheckBox
                {
                    Text = "每天",
                    Location = new Point(170, 5),
                    Size = new Size(56, 20),
                    Checked = rule == null || rule.IsEveryDay
                };
                this.Controls.Add(EveryDayCheck);

                for (int i = 0; i < 7; i++)
                {
                    // i 是"位号"（0=周一），与 ShutdownTimeRule 的位序一致。
                    // 这里刻意不碰 DayOfWeek 枚举——它的 Sunday=0 会把顺序错开一位。
                    CheckBox cb = new CheckBox
                    {
                        Text = DayTexts[i],
                        Location = new Point(232 + i * 34, 5),
                        Size = new Size(36, 20),
                        Checked = rule == null || rule.IsDayEnabled(SunFirst(i))
                    };
                    DayChecks[i] = cb;
                    this.Controls.Add(cb);
                }

                // 勾上"每天" → 七个按钮全选且禁用（它们不再是可编辑状态）
                EveryDayCheck.CheckedChanged += (s, e) => ApplyEveryDay();
                ApplyEveryDay();
            }

            /// <summary>位号（0=周一）→ DayOfWeek</summary>
            private static DayOfWeek SunFirst(int bit)
            {
                return ShutdownTimeRule.DayOfWeekForBit(bit);
            }

            private void ApplyEveryDay()
            {
                bool every = EveryDayCheck.Checked;
                for (int i = 0; i < 7; i++)
                {
                    if (every) DayChecks[i].Checked = true;
                    DayChecks[i].Enabled = !every;
                }
            }

            /// <summary>把界面状态收成一个规则对象</summary>
            public ShutdownTimeRule ToRule()
            {
                var rule = new ShutdownTimeRule
                {
                    Hour = (int)HourBox.Value,
                    Minute = (int)MinuteBox.Value
                };

                if (EveryDayCheck.Checked)
                {
                    rule.DayMask = ShutdownTimeRule.EveryDayMask;
                }
                else
                {
                    for (int i = 0; i < 7; i++)
                        if (DayChecks[i].Checked)
                            rule.SetDay(ShutdownTimeRule.DayOfWeekForBit(i), true);
                }
                return rule;
            }

            /// <summary>该行的星期选择是否为空（空规则永远不会触发，必须拦住）</summary>
            public bool HasNoDaySelected()
            {
                if (EveryDayCheck.Checked) return false;
                for (int i = 0; i < 7; i++)
                    if (DayChecks[i].Checked) return false;
                return true;
            }
        }

        // =================================================================
        // 增删行
        // =================================================================

        private void AddRuleRow(int dayMask, int hour, int minute, bool isEveryDay)
        {
            AddRuleRow(new ShutdownTimeRule
            {
                Hour = hour,
                Minute = minute,
                DayMask = isEveryDay ? ShutdownTimeRule.EveryDayMask : dayMask
            });
        }

        private void AddRuleRow(ShutdownTimeRule rule)
        {
            if (ruleRows.Count >= AppSettings.MaxShutdownRules)
            {
                MessageBox.Show(
                    "最多只能添加 " + AppSettings.MaxShutdownRules + " 个关机时间点。\n" +
                    "如需添加新的，请先删除一个已有的时间点。",
                    "已达上限", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            RuleRow row = new RuleRow(ruleRows.Count, rule);
            ruleRows.Add(row);
            rulesPanel.Controls.Add(row);
            UpdateEmptyHint();
        }

        private void RemoveLastRuleRow()
        {
            if (ruleRows.Count == 0)
            {
                MessageBox.Show("当前没有可删除的时间点。", "提示",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            RuleRow last = ruleRows[ruleRows.Count - 1];
            ruleRows.Remove(last);
            rulesPanel.Controls.Remove(last);
            last.Dispose();
            UpdateEmptyHint();
        }

        private void UpdateEmptyHint()
        {
            emptyHint.Visible = ruleRows.Count == 0;
            btnAddRule.Enabled = ruleRows.Count < AppSettings.MaxShutdownRules;
        }

        // =================================================================
        // 数据装载 / 保存
        // =================================================================

        private void LoadFromParent()
        {
            chkAutoStart.Checked = parentForm.AutoStart;
            chkSilentStart.Checked = parentForm.SilentStart;
            chkAutoStartTask.Checked = parentForm.AutoStartTask;
            chkAutoVolume.Checked = parentForm.AutoVolume;
            numVolumeLevel.Value = Clamp(parentForm.VolumeLevel, 0, 100);
            chkKillWps.Checked = parentForm.KillWps;

            // 装载现有规则。用副本——若用户的编辑被取消，
            // 运行中的配置必须保持原样（parentForm.ShutdownRules 返回的就是副本）。
            ruleRows.Clear();
            foreach (ShutdownTimeRule r in parentForm.ShutdownRules)
            {
                if (ruleRows.Count >= AppSettings.MaxShutdownRules) break;
                RuleRow row = new RuleRow(ruleRows.Count, r);
                ruleRows.Add(row);
                rulesPanel.Controls.Add(row);
            }
            UpdateEmptyHint();
        }

        private static decimal Clamp(int value, int min, int max)
        {
            if (value < min) return min;
            if (value > max) return max;
            return value;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            // ---------------- 保存前校验 ----------------
            //
            // 【为什么必须拦"一个日子都没选"】
            // 掩码为 0 的规则永远不会命中，用户设完不会有任何反馈，
            // 直到该关机的日子没关才发现。这种"静默失效"是最难排查的，
            // 宁可现在报错让他改。
            List<ShutdownTimeRule> rules = new List<ShutdownTimeRule>();
            for (int i = 0; i < ruleRows.Count; i++)
            {
                if (ruleRows[i].HasNoDaySelected())
                {
                    MessageBox.Show(
                        "第 " + (i + 1) + " 个关机时间点没有选择任何日期。\n\n" +
                        "请勾选「每天」，或至少选择一周中的某一天——\n" +
                        "否则这个时间点永远不会生效。",
                        "时间点配置不完整", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                rules.Add(ruleRows[i].ToRule());
            }

            // 检查重复（同一时刻、同一天集合）
            for (int i = 0; i < rules.Count; i++)
            {
                for (int j = i + 1; j < rules.Count; j++)
                {
                    if (rules[i].Hour == rules[j].Hour
                        && rules[i].Minute == rules[j].Minute
                        && (rules[i].DayMask & rules[j].DayMask) != 0)
                    {
                        MessageBox.Show(
                            "第 " + (i + 1) + " 个与第 " + (j + 1) +
                            " 个关机时间点在至少同一天重到了同一时刻。\n\n" +
                            "请把其中一个改成别的时间，或调整它们的生效日期。",
                            "时间点重复", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                }
            }

            // ---------------- 写回 ----------------
            parentForm.AutoStart = chkAutoStart.Checked;
            parentForm.SilentStart = chkSilentStart.Checked;
            parentForm.AutoStartTask = chkAutoStartTask.Checked;
            parentForm.AutoVolume = chkAutoVolume.Checked;
            parentForm.VolumeLevel = (int)numVolumeLevel.Value;
            parentForm.KillWps = chkKillWps.Checked;
            parentForm.ShutdownRules = rules;

            parentForm.SaveSettings();

            this.Close();
        }
    }
}
