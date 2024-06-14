using AxZuluOcx;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ZB;

namespace WindowsFormsApp1
{
    internal class DiagnosticArchive
    {
        private Panel scrollablePanel;
        private Panel scrollablePanel2;
        private Dictionary<string, Control> elementsMap;

        private TabPage tabPage;

        public DiagnosticArchive(TabPage tabPage)
        {
            this.scrollablePanel = new System.Windows.Forms.Panel();
            this.scrollablePanel.AutoScroll = true;
            this.scrollablePanel.Location = new System.Drawing.Point(0, 0);
            this.scrollablePanel.Name = "scrollablePanel";
            this.scrollablePanel.Dock = DockStyle.Left;
            this.scrollablePanel.Size = new System.Drawing.Size(795, tabPage.Size.Height);
            this.scrollablePanel.TabIndex = 0;
            tabPage.Controls.Add(this.scrollablePanel);

            this.scrollablePanel2 = new System.Windows.Forms.Panel();
            this.scrollablePanel2.AutoScroll = true;
            this.scrollablePanel2.Location = new System.Drawing.Point(800, 0);
            this.scrollablePanel2.Name = "scrollablePanel2";
            this.scrollablePanel.Dock = DockStyle.Left;
            this.scrollablePanel2.Size = new System.Drawing.Size(795, tabPage.Size.Height + 150);
            this.scrollablePanel2.TabIndex = 1;
            tabPage.Controls.Add(this.scrollablePanel2);

            this.tabPage = tabPage;
            InitializeDynamicControls();
        }

        private void InitializeDynamicControls()
        {
            elementsMap = new Dictionary<string, Control>();

            ZbDatabase database1 = new ZbDatabase();
            database1.Open(@"C:\Users\Admin\Documents\Новая папка\karta2_uch.zb");

            ZbDatabase database2 = new ZbDatabase();
            database2.Open(@"C:\Users\Admin\Documents\Новая папка\arhiv_diagnostiki.zb");

            int fieldCount1 = database1.Queries.Default.VisualQuery.Fields.Count;
            int fieldCount2 = database2.Queries.Default.VisualQuery.Fields.Count;

            CreateControlsForDatabase(database1, scrollablePanel, 1, fieldCount1);
            CreateControlsForDatabase(database2, scrollablePanel2, 2, fieldCount2);

            Button saveButton = new Button();
            saveButton.AutoSize = true;
            saveButton.Location = new System.Drawing.Point(675 - 530, fieldCount2 * 25 + 5);
            saveButton.Name = "saveButton";
            saveButton.Size = new System.Drawing.Size(44, 16);
            saveButton.TabIndex = fieldCount2 * 2 + 2;
            saveButton.Text = "Сохранить";
            saveButton.Click += new EventHandler(saveButton_Click);
            this.scrollablePanel2.Controls.Add(saveButton);

            Button entriesButton = new Button();
            entriesButton.AutoSize = true;
            entriesButton.Location = new System.Drawing.Point(675 - 530, (fieldCount2 + 1) * 25 + 5);
            entriesButton.Name = "entriesButton";
            entriesButton.Size = new System.Drawing.Size(44, 16);
            entriesButton.TabIndex = fieldCount2 * 2 + 2;
            entriesButton.Text = "Записи...";
            entriesButton.Click += new EventHandler(entriesButton_Click);
            this.scrollablePanel2.Controls.Add(entriesButton);

            ((TextBox)elementsMap["element2_Krit_dlina"]).TextChanged += new EventHandler(updateDiagRes);
        }

        void CreateControlsForDatabase(ZbDatabase database, Panel scrollablePanel, int panelIndex, int fieldCount)
        {
            for (int i = 0; i < fieldCount; i++)
            {
                if (database.Queries.Default.VisualQuery.Fields[i].Type is zbFieldType.zbftBoolean)
                {
                    CheckBox checkBox = (CheckBox)CreateControl(scrollablePanel, panelIndex, database, 0, i);
                    if (panelIndex == 2)
                    {
                        checkBox.CheckedChanged += new EventHandler(updateDiagRes);
                    }
                }
                else if (database.Queries.Default.VisualQuery.Fields[i].Name is "Raion")
                {
                    CreateControl(scrollablePanel, panelIndex, database, 1, i, bookNum: panelIndex == 1 ? 7 : 2);
                }
                else if (database.Queries.Default.VisualQuery.Fields[i].Name is "Type" && panelIndex == 2)
                {
                    ComboBox comboBox = (ComboBox)CreateControl(scrollablePanel, panelIndex, database, 1, i, bookNum: 1);
                    comboBox.SelectedIndexChanged += new EventHandler(TypeComboBox_SelectedIndexChanged);
                }
                else if (database.Queries.Default.VisualQuery.Fields[i].Name is "Diag_res" && panelIndex == 2)
                {
                    CreateControl(scrollablePanel, panelIndex, database, 1, i, bookNum: 0);
                }
                else
                {
                    CreateControl(scrollablePanel, panelIndex, database, 2, i);
                }

                var label = new Label();
                label.AutoSize = true;
                label.Location = new System.Drawing.Point(675 - 530, i * 25 + 5);
                label.Name = "label" + panelIndex + "_" + i;
                label.Size = new System.Drawing.Size(44, 16);
                label.TabIndex = i * 2 + 2;
                label.Text = database.Queries.Default.VisualQuery.Fields[i].UserName;

                scrollablePanel.Controls.Add(label);

                var fields = new[] { "Uch", "Diag_res", "Raion", "DateExpl", "Diametr", "Narabotka", "L" };
                if (panelIndex == 2)
                {
                    if (fields.Contains(database.Queries.Default.VisualQuery.Fields[i].Name))
                    {
                        elementsMap[$"element2_{database.Queries.Default.VisualQuery.Fields[i].Name}"].Enabled = false;
                    }
                    else
                    {
                        elementsMap[$"element2_{database.Queries.Default.VisualQuery.Fields[i].Name}"].Enabled = true;
                    }
                }
                else
                {
                    elementsMap[$"element1_{database.Queries.Default.VisualQuery.Fields[i].Name}"].Enabled = false;
                }
            }
        }

        string GetTextBoxValue(string key) => ((TextBox)elementsMap[key]).Text;
        bool GetCheckBoxValue(string key) => ((CheckBox)elementsMap[key]).Checked;
        int GetComboBoxIndex(string key) => ((ComboBox)elementsMap[key]).SelectedIndex;
        double ParseDouble(string text) => double.Parse(text.Replace('.', ','));
        private void saveButton_Click(object sender, EventArgs e)
        {
            try
            {
                var database = new ZbDatabase();
                database.Open(@"C:\Users\Admin\Documents\Новая папка\arhiv_diagnostiki.zb");

                var fields = new[] { "Uch", "Date", "KAO", "Dokrit", "Krit", "Krit_dlina", "Diag_res", "Comment", "Author", "Type", "Raion", "DateExpl", "Diametr", "Narabotka", "L" };
                var values = new[]
                {
                    GetTextBoxValue("element2_Uch"),
                    GetTextBoxValue("element2_Date"),
                    ParseDouble(GetTextBoxValue("element2_KAO")).ToString(),
                    GetCheckBoxValue("element2_Dokrit") ? "1" : "0",
                    GetCheckBoxValue("element2_Krit") ? "1" : "0",
                    GetTextBoxValue("element2_Krit_dlina"),
                    (GetComboBoxIndex("element2_Diag_res") + 1).ToString(),
                    GetTextBoxValue("element2_Comment"),
                    GetTextBoxValue("element2_Author"),
                    (GetComboBoxIndex("element2_Type") + 1).ToString(),
                    (GetComboBoxIndex("element2_Raion") + 1).ToString(),
                    GetTextBoxValue("element2_DateExpl"),
                    GetTextBoxValue("element2_Diametr"),
                    GetTextBoxValue("element2_Narabotka"),
                    GetTextBoxValue("element2_L")
                };
                database.AppendBaseRecord(fields, values);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}");
            }
        }

        private void entriesButton_Click(object sender, EventArgs e)
        {
            RecordsForm entriesForm = new RecordsForm();
            entriesForm.ShowDialog();
        }

        private void updateDiagRes(object sender, EventArgs e)
        {
            var dokrit = GetCheckBoxValue("element2_Dokrit");
            var krit = GetCheckBoxValue("element2_Krit");
            var krit_L = GetTextBoxValue("element2_Krit_dlina").Replace('.', ',');
            var L = GetTextBoxValue("element2_L").Replace('.', ',');
            var comboBox = (ComboBox)elementsMap["element2_Diag_res"];

            if (double.TryParse(krit_L, out double kritLength) && double.TryParse(L, out double l))
            {
                comboBox.SelectedIndex = !krit ? 0 :
                    kritLength / l <= 0.05 ? 1 :
                    kritLength / l <= 0.10 ? 2 :
                    kritLength / l <= 0.20 ? 3 : 4;
            }
            else
            {
                comboBox.SelectedIndex = 0;
            }
        }

        private void TypeComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            var typeComboBox = (ComboBox)elementsMap["element2_Type"];
            var diametrTextBox = (TextBox)elementsMap["element2_Diametr"];
            var diametrValue = typeComboBox.SelectedIndex == 0 ? elementsMap["element1_Dpod"].Text : elementsMap["element1_Dobr"].Text;

            diametrTextBox.Text = diametrValue;
        }

        private Control CreateControl(Panel scrollablePanel, int panelNum, ZbDatabase database, int type, int i, int? bookNum = null)
        {
            var element = GetControlByType(type);
            element.Location = new System.Drawing.Point(0, i * 25);
            element.Name = $"element{panelNum}_" + database.Queries.Default.VisualQuery.Fields[i].Name;
            element.Size = new System.Drawing.Size(138, 22);
            element.TabIndex = i * 2 + 1;
            elementsMap[element.Name] = element;
            scrollablePanel.Controls.Add(element);

            if (type == 1)
            {
                PopulateComboBox(element as ComboBox, (int)bookNum, database);
                (element as ComboBox).SelectedIndex = 0;
            }

            return element;
        }

        private Control GetControlByType(int type)
        {
            switch (type)
            {
                case 0:
                    return new CheckBox();
                case 1:
                    return new ComboBox() { DropDownStyle = ComboBoxStyle.DropDownList };
                default:
                    return new TextBox();
            }
        }

        private void PopulateComboBox(ComboBox comboBox, int bookNum, ZbDatabase database)
        {
            for (int j = 1; ; j++)
            {
                Debug.WriteLine(database.Books[bookNum].SimpleBook.FindValue(j.ToString(), out string comboboxValue));
                if (comboboxValue is null || comboboxValue == string.Empty) break;
                comboBox.Items.Add(comboboxValue);
            }
        }

        internal void ObjectSelect(ZuluLib.Layer activeLayer)
        {
            var SysSelected = activeLayer.CurrentID;
            Debug.WriteLine(SysSelected);
            var activElem = activeLayer.Elements.GetElement(SysSelected);
            var objectType = activElem.get_Type();
            if (objectType is null || objectType.Name != "Участки") return;

            ZbDatabase database = new ZbDatabase();
            database.Open(@"C:\Users\Admin\Documents\Новая папка\karta2_uch.zb");


            int fieldCount1 = database.Queries.Default.VisualQuery.Fields.Count;
            for (int i = 0; i < fieldCount1; i++)
            {
                for (int j = 1; j <= 2; j++)
                {
                    var fieldName = $"element{j}_" + database.Queries.Default.VisualQuery.Fields[i].Name;
                    var fieldValue = database.SelectByKey(SysSelected).FieldValue[i + 1].ToString();
                    Debug.WriteLine(fieldValue);

                    if (elementsMap.ContainsKey(fieldName))
                    {
                        if (elementsMap[fieldName] is ComboBox)
                        {
                            ((ComboBox)elementsMap[fieldName]).SelectedIndex = int.Parse(fieldValue) - 1;
                        }
                        else
                        {
                            elementsMap[fieldName].Text = fieldValue;
                        }
                    }
                }

            }

            int y = ((ComboBox)elementsMap["element2_Type"]).SelectedIndex;
            if (y == 0)
            {
                elementsMap["element2_Diametr"].Text = elementsMap["element1_Dpod"].Text;
            }
            else
            {
                elementsMap["element2_Diametr"].Text = elementsMap["element1_Dobr"].Text;
            }
        }
    }


}
