using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using SapfirLib;

namespace AcSapfir
{
    public class StoreyGeneratorForm : Form
    {
        private DataGridView dgv;
        private Button btnAdd, btnDel, btnUp, btnDown;
        private Button btnTplBasement, btnTplFirst, btnTplTypical, btnTplAttic;
        private NumericUpDown nudTypicalHeight, nudTypicalCount;
        private Button btnCreate, btnCancel;
        private Label lblStatus;
        private SapfirLib.Application _app;
        private SapfirDoc _doc;
        private int _typicalCounter = 1;

        public List<Tuple<string, double>> Result { get; private set; }

        public StoreyGeneratorForm(SapfirLib.Application app, SapfirDoc doc)
        {
            _app = app;
            _doc = doc;
            Result = null;
            BuildUI();
            UpdateStatus();
        }

        void BuildUI()
        {
            Text = "Генератор этажей";
            Size = new Size(620, 520);
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;

            var topPanel = new FlowLayoutPanel
            {
                Location = new Point(10, 10),
                Size = new Size(584, 30),
                FlowDirection = FlowDirection.LeftToRight
            };

            btnAdd = new Button { Text = "Добавить", Size = new Size(75, 25) };
            btnDel = new Button { Text = "Удалить", Size = new Size(75, 25) };
            btnUp = new Button { Text = "▲", Size = new Size(35, 25) };
            btnDown = new Button { Text = "▼", Size = new Size(35, 25) };

            btnAdd.Click += (s, e) => AddRow();
            btnDel.Click += (s, e) => DelRow();
            btnUp.Click += (s, e) => MoveRow(-1);
            btnDown.Click += (s, e) => MoveRow(1);

            topPanel.Controls.AddRange(new Control[] { btnAdd, btnDel, btnUp, btnDown });
            Controls.Add(topPanel);

            dgv = new DataGridView
            {
                Location = new Point(10, 45),
                Size = new Size(584, 200),
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeRows = false,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            };
            dgv.Columns.Add("Name", "Имя этажа");
            dgv.Columns["Name"].FillWeight = 60;
            var elevCol = new DataGridViewTextBoxColumn
            {
                Name = "Elevation",
                HeaderText = "Отметка, м",
                FillWeight = 40
            };
            dgv.Columns.Add(elevCol);
            Controls.Add(dgv);

            var tplPanel = new FlowLayoutPanel
            {
                Location = new Point(10, 250),
                Size = new Size(584, 55),
                FlowDirection = FlowDirection.LeftToRight
            };

            btnTplBasement = new Button { Text = "Подвал", Size = new Size(75, 25) };
            btnTplFirst = new Button { Text = "1-й этаж", Size = new Size(75, 25) };
            btnTplTypical = new Button { Text = "Типовые ×", Size = new Size(75, 25) };
            nudTypicalCount = new NumericUpDown { Minimum = 1, Maximum = 99, Value = 3, Size = new Size(45, 25) };
            var lblH = new Label { Text = "  Высота этажа:", Size = new Size(90, 25), TextAlign = ContentAlignment.MiddleRight };
            nudTypicalHeight = new NumericUpDown
            {
                Minimum = 0.5m, Maximum = 20, Value = 3.0m,
                DecimalPlaces = 2, Increment = 0.1m, Size = new Size(60, 25)
            };
            var lblM = new Label { Text = "м", Size = new Size(20, 25), TextAlign = ContentAlignment.MiddleLeft };
            btnTplAttic = new Button { Text = "Чердак", Size = new Size(75, 25) };

            btnTplBasement.Click += (s, e) => AddTemplate("Подвал", -3.0);
            btnTplFirst.Click += (s, e) => AddTemplate("1 этаж", 0.0);
            btnTplTypical.Click += (s, e) => AddTypicalFloors();
            btnTplAttic.Click += (s, e) => AddTemplate("Чердак", GetTopElevation() + (double)nudTypicalHeight.Value);

            tplPanel.Controls.AddRange(new Control[] {
                new Label { Text = "Шаблоны: ", Size = new Size(65, 25), TextAlign = ContentAlignment.MiddleRight },
                btnTplBasement, btnTplFirst, btnTplTypical, nudTypicalCount,
                lblH, nudTypicalHeight, lblM, btnTplAttic
            });
            Controls.Add(tplPanel);

            lblStatus = new Label
            {
                Location = new Point(10, 315),
                Size = new Size(584, 25),
                TextAlign = ContentAlignment.MiddleLeft
            };
            Controls.Add(lblStatus);

            var bottomPanel = new FlowLayoutPanel
            {
                Location = new Point(10, 350),
                Size = new Size(584, 35),
                FlowDirection = FlowDirection.RightToLeft
            };

            btnCancel = new Button { Text = "Отмена", Size = new Size(100, 30) };
            btnCreate = new Button { Text = "Создать этажи", Size = new Size(130, 30) };

            btnCreate.Click += (s, e) => OnCreate();
            btnCancel.Click += (s, e) => Close();

            bottomPanel.Controls.AddRange(new Control[] { btnCancel, btnCreate });
            Controls.Add(bottomPanel);
        }

        void AddRow(string name = "", double elev = 0)
        {
            var nextName = name;
            if (string.IsNullOrEmpty(nextName))
            {
                int idx = dgv.Rows.Count + 1;
                nextName = idx + " этаж";
            }
            dgv.Rows.Add(nextName, elev.ToString("F2"));
            dgv.CurrentCell = dgv.Rows[dgv.Rows.Count - 1].Cells[0];
        }

        void DelRow()
        {
            if (dgv.SelectedRows.Count > 0)
                dgv.Rows.Remove(dgv.SelectedRows[0]);
        }

        void MoveRow(int dir)
        {
            if (dgv.SelectedRows.Count == 0) return;
            int idx = dgv.SelectedRows[0].Index;
            int newIdx = idx + dir;
            if (newIdx < 0 || newIdx >= dgv.Rows.Count) return;

            var name = dgv.Rows[idx].Cells[0].Value?.ToString() ?? "";
            var elev = dgv.Rows[idx].Cells[1].Value?.ToString() ?? "0";
            dgv.Rows.RemoveAt(idx);
            dgv.Rows.Insert(newIdx, name, elev);
            dgv.CurrentCell = dgv.Rows[newIdx].Cells[0];
            dgv.Rows[newIdx].Selected = true;
        }

        void AddTemplate(string name, double elev)
        {
            AddRow(name, elev);
        }

        void AddTypicalFloors()
        {
            int count = (int)nudTypicalCount.Value;
            double h = (double)nudTypicalHeight.Value;
            double baseElev = GetTopElevation();

            for (int i = 0; i < count; i++)
            {
                int num = _typicalCounter++;
                string name;
                if (num == 1) name = "1 этаж";
                else if (num == 2) name = "2 этаж";
                else if (num == 3) name = "3 этаж";
                else name = num + " этаж";
                AddRow(name, baseElev + i * h);
            }
        }

        double GetTopElevation()
        {
            double top = 0;
            foreach (DataGridViewRow row in dgv.Rows)
            {
                if (row.Cells[1].Value != null &&
                    double.TryParse(row.Cells[1].Value.ToString(),
                        System.Globalization.NumberStyles.Float,
                        System.Globalization.CultureInfo.InvariantCulture,
                        out double elev))
                {
                    if (elev > top) top = elev;
                }
            }
            return top;
        }

        void UpdateStatus()
        {
            try
            {
                var view = _app.GetActiveSapfirView();
                SapfirDoc d = view != null ? view.GetDocument() : _app.GetActiveDoc();
                if (d == null) { lblStatus.Text = "Sapfir: нет активного документа"; return; }
                var proj = d.CountProjects > 0 ? d.GetActiveProject() : null;
                if (proj == null) lblStatus.Text = "Sapfir: " + d.Title + " - нет проектов";
                else lblStatus.Text = "Sapfir: " + d.Title + " / Проект ID=" + proj.ID + " / этажей: " + proj.CountStorey;
            }
            catch
            {
                lblStatus.Text = "Sapfir: недоступен";
            }
        }

        void OnCreate()
        {
            var data = new List<Tuple<string, double>>();
            foreach (DataGridViewRow row in dgv.Rows)
            {
                if (row.IsNewRow) continue;
                string name = row.Cells[0].Value?.ToString()?.Trim();
                if (string.IsNullOrEmpty(name)) continue;
                if (!double.TryParse(row.Cells[1].Value?.ToString(),
                        System.Globalization.NumberStyles.Float,
                        System.Globalization.CultureInfo.InvariantCulture,
                        out double elev)) continue;
                data.Add(Tuple.Create(name, elev));
            }

            if (data.Count == 0)
            {
                MessageBox.Show("Добавьте хотя бы один этаж.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Result = data.OrderBy(item => item.Item2).ToList();
            DialogResult = DialogResult.OK;
            Close();
        }
    }
}
