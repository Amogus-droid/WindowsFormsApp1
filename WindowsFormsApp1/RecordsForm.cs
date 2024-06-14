using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ZB;

namespace WindowsFormsApp1
{
    public partial class RecordsForm : Form
    {
        Label[] labels;
        public RecordsForm()
        {
            InitializeComponent();

            ZbDatabase database = new ZbDatabase();
            database.Open(@"C:\Users\Admin\Documents\Новая папка\arhiv_diagnostiki.zb");
            int fieldCount = database.Queries.Default.VisualQuery.Fields.Count;

            this.AutoSize = true;

            var combobox = new ComboBox() { DropDownStyle = ComboBoxStyle.DropDownList };
            combobox.AutoSize = true;
            combobox.Location = new System.Drawing.Point(0, 5);
            combobox.Name = "combobox";
            combobox.Size = new System.Drawing.Size(44, 16);
            combobox.TabIndex = 0;
            combobox.SelectedIndexChanged += new EventHandler(FillLabels);
            this.Controls.Add(combobox);

            var records = database.SelectAll();
            int count = 0;
            if (records.MoveFirst()) count++;
            while (records.MoveNext()) count++;
            records.MoveFirst();

            for (int i = 1; i <= count; i++)
            {
                combobox.Items.Add(i.ToString());
            }

            int maxLength = 0;
            for (int i = 1; i < fieldCount + 1; i++)
            {
                var label = new System.Windows.Forms.Label();
                label.AutoSize = true;
                label.Location = new System.Drawing.Point(0, i * 25 + 5);
                label.Name = "label" + "_" + i;
                label.Size = new System.Drawing.Size(44, 16);
                label.TabIndex = i * 2 + 2;
                label.Text = database.Queries.Default.VisualQuery.Fields[i-1].UserName;
                this.Controls.Add(label);

                if (maxLength < label.Bounds.Width) maxLength = label.Bounds.Width;
            }

            labels = new Label[fieldCount];
            for (int i = 1; i < fieldCount + 1; i++)
            {
                var label = new System.Windows.Forms.Label();
                label.AutoSize = true;
                label.Location = new System.Drawing.Point(maxLength, i * 25 + 5);
                label.Name = "label" + "_" + i;
                label.Size = new System.Drawing.Size(44, 16);
                label.TabIndex = i * 2 + 2;
                //label.Text = database.Queries.Default.VisualQuery.Fields[i-1].UserName;
                this.Controls.Add(label);
                labels[i-1] = label;
            }

            this.Width *= 2;
        }

        private void FillLabels(object sender, EventArgs e)
        {
            int selectedRecord = ((ComboBox)sender).SelectedIndex;

            ZbDatabase database = new ZbDatabase();
            database.Open(@"C:\Users\Admin\Documents\Новая папка\arhiv_diagnostiki.zb");
            int fieldCount = database.Queries.Default.VisualQuery.Fields.Count;
            var records = database.SelectAll();

            while (selectedRecord > 0)
            {
                selectedRecord--;
                records.MoveNext();
            }

            for (int i = 1; i <= fieldCount; i++)
            {
                labels[i - 1].Text = records.FieldDisplayValue[i];
            }
        }
    }
}
