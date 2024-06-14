using ClosedXML.Excel;
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
    internal class HeatConsumptionSystemWaterVolumeCalculator
    {
        private TabPage tabPage;
        private HashSet<int> sysHashSet = new HashSet<int>();
        private Dictionary<string, Control> elementsMap = new Dictionary<string, Control>();

        public HeatConsumptionSystemWaterVolumeCalculator(TabPage tabPage)
        {
            this.tabPage = tabPage;

            Label label = new Label();
            label.AutoSize = true;
            label.Name = "label";
            label.Size = new System.Drawing.Size(44, 16);
            label.Text = "Выбрано элементов: 0";
            tabPage.Controls.Add(label);
            elementsMap["label"] = label;

            Button button = new Button();
            button.AutoSize = true;
            button.Location = new System.Drawing.Point(0, 25);
            button.Name = "saveButton";
            button.Size = new System.Drawing.Size(44, 16);
            button.Text = "Получить отчет";
            button.Click += new EventHandler(saveButton_Click);
            tabPage.Controls.Add(button);
            elementsMap["saveButton"] = label;
        }
        private void saveButton_Click(object sender, EventArgs e)
        {
            ZbDatabase database = new ZbDatabase();
            database.Open(@"C:\Users\Admin\Documents\Новая папка\karta2_uch.zb");

            double sumOtopl = 0;
            double sumVent = 0;
            foreach (var SysSelected in sysHashSet)
            {
                var obj = database.SelectByKey(SysSelected);
                sumOtopl += double.Parse(database.SelectByKey(SysSelected).FieldValue[obj.GetFieldIndexByName(zbNameType.zbShortName, "Otopl")].Replace('.', ','));
                sumVent += double.Parse(database.SelectByKey(SysSelected).FieldValue[obj.GetFieldIndexByName(zbNameType.zbShortName, "Vent")].Replace('.', ','));
            }

            string filePath = "output.xlsx";

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Sheet1");

                worksheet.Rows().Height = 0.17 * 72;

                var headerCells = worksheet.Range("A1:E8");
                headerCells.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                headerCells.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                headerCells.Style.Alignment.WrapText = true;
                headerCells.Style.Font.FontName = "Times New Roman";
                headerCells.Style.Font.FontSize = 12;

                headerCells = worksheet.Range("A1:E7");
                headerCells.Style.Border.TopBorder = XLBorderStyleValues.Thin;
                headerCells.Style.Border.TopBorderColor = XLColor.Black;
                headerCells.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                headerCells.Style.Border.LeftBorderColor = XLColor.Black;
                headerCells.Style.Border.RightBorder = XLBorderStyleValues.Thin;
                headerCells.Style.Border.RightBorderColor = XLColor.Black;
                headerCells.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                headerCells.Style.Border.BottomBorderColor = XLColor.Black;

                worksheet.Columns().Width = 15;
                worksheet.Column(1).Width = 4.5 * 72 / 6;
                worksheet.Row(1).Height = 1.36 * 72;

                worksheet.Rows(3, 8).Height = 0.22 * 72;

                worksheet.Range("D11:D15").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Range("A11:A15").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                worksheet.Cell("B11").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

                worksheet.Range("D12:D15").Style.Font.Bold = true;
                worksheet.Range("B12:B14").Style.Font.Bold = true;
                worksheet.Range("D8:E8").Style.Font.Bold = true;
                worksheet.Cell("A12").Style.Font.Bold = true;
                worksheet.Cell("A15").Style.Font.Bold = true;

                worksheet.Cell("A1").Value = "Теплопотребляющее оборудование";
                worksheet.Cell("B1").Value = "Температурный график";
                worksheet.Cell("C1").Value = "Тепловая нагрузка Qов МВт";
                worksheet.Cell("D1").Value = "Удельный объем воды в системе, м3/МВт";
                worksheet.Cell("E1").Value = "Объем воды в системе Vnотр, м3";
                worksheet.Cell("A3").Value = "Водяные системы теплоснабжения";
                worksheet.Cell("A4").Value = "Радиаторы чугунные высотой 500 мм";
                worksheet.Cell("A5").Value = "Радиаторы стальные панельные высотой 500 мм";
                worksheet.Cell("A6").Value = "Регистры из стальных труб";
                worksheet.Cell("A7").Value = "Калориферные отопительно-вентиляционные агрегаты";
                worksheet.Cell("B4").Value = "95-70";
                worksheet.Cell("B5").Value = "95-70";
                worksheet.Cell("B6").Value = "95-70";
                worksheet.Cell("B7").Value = "110-70";
                worksheet.Cell("D4").Value = "16.8";
                worksheet.Cell("D5").Value = "10.1";
                worksheet.Cell("D6").Value = "31.8";
                worksheet.Cell("D7").Value = "6.4";
                worksheet.Cell("D8").Value = "Итого:";

                worksheet.Cell("A11").Value = "Тепловая нагрузка отопление Гкал/ч";
                worksheet.Cell("A14").Value = "Тепловая нагрузка на вентиляцию Гкал/ч";
                worksheet.Cell("B11").Value = "%";
                worksheet.Cell("B12").Value = "79";
                worksheet.Cell("B13").Value = "16";
                worksheet.Cell("B14").Value = "5";
                worksheet.Cell("D11").Value = "Перевод МВт";

                worksheet.Range("A1:A2").Merge();
                worksheet.Range("B1:B2").Merge();
                worksheet.Range("D1:D2").Merge();
                worksheet.Range("A3:E3").Merge();

                worksheet.Cell("A12").Value = sumOtopl;
                worksheet.Cell("A15").Value = sumVent;

                worksheet.Cell("D12").FormulaA1 = "A12 * 1.163";
                worksheet.Cell("D15").FormulaA1 = "A15 * 1.163";

                worksheet.Cell("C12").FormulaA1 = "D12 * 0.01 * B12";
                worksheet.Cell("C13").FormulaA1 = "D12 * 0.01 * B13";
                worksheet.Cell("C14").FormulaA1 = "D12 * 0.01 * B14";

                worksheet.Cell("C4").FormulaA1 = "C12";
                worksheet.Cell("C5").FormulaA1 = "C13";
                worksheet.Cell("C6").FormulaA1 = "C14";
                worksheet.Cell("C7").FormulaA1 = "D15";

                worksheet.Cell("E4").FormulaA1 = "0.3 * C4 * D4";
                worksheet.Cell("E5").FormulaA1 = "0.3 * C5 * D5";
                worksheet.Cell("E6").FormulaA1 = "0.3 * C6 * D6";
                worksheet.Cell("E7").FormulaA1 = "0.3 * C7 * D7";
                worksheet.Cell("E8").FormulaA1 = "E4 + E5 + E6 + E7";

                workbook.SaveAs(filePath);
            }

            Process.Start("output.xlsx");
        }

        internal void ObjectSelect(ZuluLib.Layer activeLayer)
        {
            var sysSelected = activeLayer.CurrentID;

            if (sysHashSet.Contains(sysSelected))
            {
                sysHashSet.Remove(sysSelected);
            }
            else
            {
                var activElem = activeLayer.Elements.GetElement(sysSelected);
                var objectType = activElem.get_Type();
                if (objectType is null || objectType.Name != "Участки") return;

                sysHashSet.Add(sysSelected);
            }

            elementsMap["label"].Text = $"Выбрано элементов: {sysHashSet.Count}";
        }
    }
}
