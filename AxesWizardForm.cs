using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Autodesk.AutoCAD.DatabaseServices;
using AcLine = Autodesk.AutoCAD.DatabaseServices.Line;

namespace AcSapfir
{
    public class AxesWizardForm : Form
    {
        private Label lblInfo;
        private Panel pnlGroups;
        private Button btnCreate, btnCancel;

        private static readonly char[] CyrillicLetters =
            "АБВГДЕЖИКЛМНОПРСТУФХЦЧШЩЫЭЮЯ".ToCharArray();

        private class LineGroup
        {
            public double DirX, DirY;
            public List<AcLine> Lines = new List<AcLine>();
            public ComboBox cmbRule;
            public ComboBox cmbStartLetter;
            public NumericUpDown nudStartNum, nudStep;
            public TextBox txtPrefix, txtTemplate;
            public Label lblPreview;
            public int Rule
            {
                get { return cmbRule.SelectedIndex; }
            }
        }

        private List<LineGroup> groups = new List<LineGroup>();

        public List<Tuple<AcLine, string>> Result { get; private set; }

        public AxesWizardForm(List<AcLine> lines)
        {
            Result = null;
            BuildUI();
            GroupLines(lines);
            PopulateGroups();
        }

        void BuildUI()
        {
            Text = "Мастер координационных осей";
            Size = new Size(680, 550);
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;

            lblInfo = new Label
            {
                Location = new Point(10, 10),
                Size = new Size(644, 25),
                TextAlign = ContentAlignment.MiddleLeft
            };
            Controls.Add(lblInfo);

            pnlGroups = new Panel
            {
                Location = new Point(10, 40),
                Size = new Size(644, 410),
                AutoScroll = true,
                BorderStyle = BorderStyle.FixedSingle
            };
            Controls.Add(pnlGroups);

            btnCancel = new Button { Text = "Отмена", Location = new Point(464, 460), Size = new Size(100, 30) };
            btnCreate = new Button { Text = "Создать оси", Location = new Point(570, 460), Size = new Size(100, 30) };

            btnCreate.Click += (s, e) => OnCreate();
            btnCancel.Click += (s, e) => Close();

            Controls.AddRange(new Control[] { btnCancel, btnCreate });
        }

        void GroupLines(List<AcLine> lines)
        {
            groups.Clear();
            var remaining = new List<AcLine>(lines);

            while (remaining.Count > 0)
            {
                var line = remaining[0];
                remaining.RemoveAt(0);
                double dx = line.EndPoint.X - line.StartPoint.X;
                double dy = line.EndPoint.Y - line.StartPoint.Y;
                double len = Math.Sqrt(dx * dx + dy * dy);
                if (len < 1e-9) continue;
                double nx = dx / len;
                double ny = dy / len;

                var group = new LineGroup { DirX = nx, DirY = ny };
                group.Lines.Add(line);

                for (int i = remaining.Count - 1; i >= 0; i--)
                {
                    var other = remaining[i];
                    double odx = other.EndPoint.X - other.StartPoint.X;
                    double ody = other.EndPoint.Y - other.StartPoint.Y;
                    double olen = Math.Sqrt(odx * odx + ody * ody);
                    if (olen < 1e-9) { remaining.RemoveAt(i); continue; }
                    double onx = odx / olen;
                    double ony = ody / olen;

                    double dot = Math.Abs(nx * onx + ny * ony);
                    if (dot > 0.999)
                    {
                        group.Lines.Add(other);
                        remaining.RemoveAt(i);
                    }
                }

                double perpX = -ny, perpY = nx;
                group.Lines.Sort((a, b) =>
                {
                    double ca = (a.StartPoint.X + a.EndPoint.X) * 0.5;
                    double cb = (b.StartPoint.X + b.EndPoint.X) * 0.5;
                    double pa = ca * perpX + (a.StartPoint.Y + a.EndPoint.Y) * 0.5 * perpY;
                    double pb = cb * perpX + (b.StartPoint.Y + b.EndPoint.Y) * 0.5 * perpY;
                    return pa.CompareTo(pb);
                });

                groups.Add(group);
            }
        }

        void PopulateGroups()
        {
            lblInfo.Text = "Линий: " + groups.Sum(g => g.Lines.Count) + ", групп: " + groups.Count;
            pnlGroups.Controls.Clear();
            int y = 5;

            for (int gi = 0; gi < groups.Count; gi++)
            {
                var group = groups[gi];
                int count = group.Lines.Count;

                var gbox = new GroupBox
                {
                    Text = "Группа " + (gi + 1) + " (" + count + " линий, вектор " + group.DirX.ToString("F4") + "," + group.DirY.ToString("F4") + ")",
                    Location = new Point(5, y),
                    Size = new Size(615, 120),
                    Tag = group
                };

                int cy = 20;
                var lblRule = new Label { Text = "Правило:", Location = new Point(10, cy), Size = new Size(55, 25), TextAlign = ContentAlignment.MiddleRight };
                var cmbRule = new ComboBox
                {
                    Location = new Point(70, cy),
                    Size = new Size(160, 25),
                    DropDownStyle = ComboBoxStyle.DropDownList,
                    Items = { "Буквы кириллицей", "Цифры", "Буквы + цифры", "Шаблон" },
                    SelectedIndex = 0
                };
                group.cmbRule = cmbRule;

                var lblStart = new Label { Text = "Начало:", Location = new Point(240, cy), Size = new Size(55, 25), TextAlign = ContentAlignment.MiddleRight };

                var cmbLetter = new ComboBox
                {
                    Location = new Point(300, cy),
                    Size = new Size(60, 25),
                    DropDownStyle = ComboBoxStyle.DropDownList
                };
                foreach (char c in CyrillicLetters) cmbLetter.Items.Add(c.ToString());
                cmbLetter.SelectedIndex = 0;
                group.cmbStartLetter = cmbLetter;

                var nudNum = new NumericUpDown
                {
                    Location = new Point(300, cy),
                    Size = new Size(60, 25),
                    Minimum = 1, Maximum = 999, Value = 1,
                    Visible = false
                };
                group.nudStartNum = nudNum;

                var lblStep = new Label { Text = "Шаг:", Location = new Point(370, cy), Size = new Size(40, 25), TextAlign = ContentAlignment.MiddleRight };
                var nudStep = new NumericUpDown
                {
                    Location = new Point(415, cy),
                    Size = new Size(50, 25),
                    Minimum = 1, Maximum = 100, Value = 1,
                    Visible = true
                };
                group.nudStep = nudStep;

                var lblPrefix = new Label { Text = "Префикс буквы:", Location = new Point(240, cy), Size = new Size(95, 25), TextAlign = ContentAlignment.MiddleRight, Visible = false };
                var txtPrefix = new TextBox { Location = new Point(340, cy), Size = new Size(80, 25), Text = "А", Visible = false };
                group.txtPrefix = txtPrefix;

                var lblTemplate = new Label { Text = "Шаблон:", Location = new Point(240, cy), Size = new Size(55, 25), TextAlign = ContentAlignment.MiddleRight, Visible = false };
                var txtTemplate = new TextBox { Location = new Point(300, cy), Size = new Size(200, 25), Text = "Ось {0}", Visible = false };
                txtTemplate.TextChanged += (s, e) => UpdatePreview(group);
                group.txtTemplate = txtTemplate;

                var lblPrev = new Label
                {
                    Location = new Point(10, cy + 30),
                    Size = new Size(590, 25),
                    TextAlign = ContentAlignment.MiddleLeft,
                    ForeColor = Color.DarkGreen
                };
                group.lblPreview = lblPrev;

                cmbRule.SelectedIndexChanged += (s, e) =>
                {
                    int rule = cmbRule.SelectedIndex;
                    cmbLetter.Visible = (rule == 0 || rule == 3);
                    nudNum.Visible = (rule == 1 || rule == 2 || rule == 3);
                    nudStep.Visible = (rule == 1);
                    lblPrefix.Visible = txtPrefix.Visible = false;
                    lblTemplate.Visible = txtTemplate.Visible = (rule == 3);
                    lblStart.Visible = (rule == 0 || rule == 3);
                    lblStep.Visible = (rule == 1);
                    UpdatePreview(group);
                };

                gbox.Controls.AddRange(new Control[] { lblRule, cmbRule, lblStart, cmbLetter, nudNum, lblStep, nudStep, lblPrefix, txtPrefix, lblTemplate, txtTemplate, lblPrev });
                pnlGroups.Controls.Add(gbox);

                cmbRule_InitialState(group);

                y += 130;
            }
        }

        void cmbRule_InitialState(LineGroup group)
        {
            int rule = group.Rule;
            group.cmbStartLetter.Visible = (rule == 0 || rule == 3);
            group.nudStartNum.Visible = (rule == 1 || rule == 2 || rule == 3);
            group.nudStep.Visible = (rule == 1);
            group.txtPrefix.Visible = false;
            group.txtTemplate.Visible = (rule == 3);
            UpdatePreview(group);
        }

        List<string> GenerateNames(LineGroup group)
        {
            int rule = group.Rule;
            int count = group.Lines.Count;
            var names = new List<string>();

            if (rule == 0)
            {
                int startIdx = group.cmbStartLetter.SelectedIndex;
                if (startIdx < 0) startIdx = 0;
                for (int i = 0; i < count; i++)
                {
                    int idx = startIdx + i;
                    if (idx >= CyrillicLetters.Length) break;
                    names.Add(CyrillicLetters[idx].ToString());
                }
            }
            else if (rule == 1)
            {
                int start = (int)group.nudStartNum.Value;
                int step = (int)group.nudStep.Value;
                for (int i = 0; i < count; i++)
                    names.Add((start + i * step).ToString());
            }
            else if (rule == 2)
            {
                int startNum = (int)group.nudStartNum.Value;
                for (int i = 0; i < count; i++)
                {
                    int idx = i % CyrillicLetters.Length;
                    names.Add(CyrillicLetters[idx] + startNum.ToString());
                }
            }
            else if (rule == 3)
            {
                string tmpl = group.txtTemplate.Text;
                if (string.IsNullOrEmpty(tmpl)) tmpl = "Ось {L}";
                int letIdx = group.cmbStartLetter.SelectedIndex >= 0 ? group.cmbStartLetter.SelectedIndex : 0;
                int numVal = (int)group.nudStartNum.Value;
                for (int i = 0; i < count; i++)
                {
                    string name = tmpl;
                    name = name.Replace("{L}", CyrillicLetters[Math.Min(letIdx + i, CyrillicLetters.Length - 1)].ToString());
                    name = name.Replace("{N}", (numVal + i).ToString());
                    names.Add(name);
                }
            }
            return names;
        }

        void UpdatePreview(LineGroup group)
        {
            var names = GenerateNames(group);
            group.lblPreview.Text = "→ " + string.Join(", ", names);
        }

        void OnCreate()
        {
            Result = new List<Tuple<AcLine, string>>();
            foreach (var group in groups)
            {
                var names = GenerateNames(group);
                for (int i = 0; i < group.Lines.Count && i < names.Count; i++)
                    Result.Add(Tuple.Create(group.Lines[i], names[i]));
            }

            if (Result.Count == 0)
            {
                MessageBox.Show("Нет осей для создания.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            DialogResult = DialogResult.OK;
            Close();
        }
    }
}
