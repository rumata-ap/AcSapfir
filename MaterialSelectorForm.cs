using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using SapfirLib;

namespace AcSapfir
{
    public class MaterialSelectorForm : Form
    {
        private ListBox lstMaterials;
        private Button btnOk, btnCancel;
        private Label lblInfo;

        public string SelectedGuid { get; private set; }
        public string SelectedName { get; private set; }

        public MaterialSelectorForm(SapfirDoc doc, string elementType)
        {
            SelectedGuid = null;
            SelectedName = null;

            Text = "Выбор материала для: " + elementType;
            Size = new Size(420, 400);
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;

            lblInfo = new Label
            {
                Location = new Point(10, 10),
                Size = new Size(384, 30),
                Text = "Загрузка материалов..."
            };
            Controls.Add(lblInfo);

            lstMaterials = new ListBox
            {
                Location = new Point(10, 45),
                Size = new Size(384, 260),
                DisplayMember = "Name"
            };
            Controls.Add(lstMaterials);

            btnCancel = new Button { Text = "Пропустить (по умолчанию)", Location = new Point(10, 320), Size = new Size(180, 30) };
            btnOk = new Button { Text = "OK", Location = new Point(214, 320), Size = new Size(180, 30), Enabled = false };

            btnOk.Click += (s, e) =>
            {
                if (lstMaterials.SelectedItem is MaterialItem item)
                {
                    SelectedGuid = item.Guid;
                    SelectedName = item.Name;
                }
                DialogResult = DialogResult.OK;
                Close();
            };
            btnCancel.Click += (s, e) =>
            {
                DialogResult = DialogResult.Cancel;
                Close();
            };

            lstMaterials.SelectedIndexChanged += (s, e) => btnOk.Enabled = lstMaterials.SelectedIndex >= 0;
            lstMaterials.DoubleClick += (s, e) => { if (btnOk.Enabled) btnOk.PerformClick(); };

            Controls.AddRange(new Control[] { btnCancel, btnOk });

            LoadMaterials(doc);
        }

    void LoadMaterials(SapfirDoc doc)
    {
        var items = new List<MaterialItem>();
        try
        {
            for (int di = 1; di <= 100; di++)
            {
                SapfirDoc d = null;
                try
                {
                    var app = doc.GetApplication();
                    if (app == null) break;
                    d = app.GetDocByID(di);
                }
                catch { break; }
                if (d == null) break;

                for (int pi = 0; pi < d.CountProjects; pi++)
                {
                    var proj = d.GetProjectByIndex(pi);
                    for (int si = 0; si < proj.CountStorey; si++)
                    {
                        var storey = proj.GetStoreyByIndex(si);
                        for (int mi = 0; mi < storey.CountModel; mi++)
                        {
                            var model = storey.GetModelByIndex(mi);
                            int mt = model.TypeModel;
                            if (mt != 251 && mt != 252) continue;
                            try
                            {
                                string guid = "";
                                string name = "";
                                try { guid = model.Parameter["M_GUID"]?.ToString(); } catch { }
                                try { name = model.Parameter["M_NAME"]?.ToString(); } catch { }
                                if (!string.IsNullOrEmpty(guid) && !items.Exists(x => x.Guid == guid))
                                    items.Add(new MaterialItem { Guid = guid, Name = name ?? guid });
                            }
                            catch { }
                        }
                    }
                }
            }
        }
        catch { }

        if (items.Count == 0)
        {
            var hardcoded = new Dictionary<string, string>
            {
                {"ab7d2fa2-df44-4d8f-8d46-71b7fb919bde", "Бетон тяжёлый (стена)"},
                {"79f404bd-50e9-4058-b664-c159ffdd0ce8", "Бетон тяжёлый (плита)"},
                {"239d15bb-4253-4546-89f1-2d90300a3c79", "Бетон тяжёлый (фундамент)"},
                {"bf0bce6b-c444-4d27-9b90-d0a38b6870b7", "Железобетон (колонна)"},
                {"25304e6e-8203-4968-b6fb-22b82250b55c", "Железобетон балок"}
            };
            foreach (var kvp in hardcoded)
                items.Add(new MaterialItem { Guid = kvp.Key, Name = kvp.Value });
            lblInfo.Text = "Встроенные материалы (" + items.Count + ")";
        }
        else
        {
            lblInfo.Text = "Материалов: " + items.Count;
        }

        lstMaterials.DataSource = items;
    }
    }

    public class MaterialItem
    {
        public string Guid { get; set; }
        public string Name { get; set; }
    }
}
