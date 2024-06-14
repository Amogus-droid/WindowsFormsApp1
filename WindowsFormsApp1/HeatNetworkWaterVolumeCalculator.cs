using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office2019.Excel.RichData2;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ZB;
using ZuluLib;

namespace WindowsFormsApp1
{
    internal class HeatNetworkWaterVolumeCalculator
    {
        private TabPage tabPage;
        private HashSet<int> sysHashSet = new HashSet<int>();
        private Dictionary<string, Control> elementsMap = new Dictionary<string, Control>();

        public HeatNetworkWaterVolumeCalculator(TabPage tabPage)
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
            List<HeatNetworkWaterVolumeCalculatorElement> elements = new List<HeatNetworkWaterVolumeCalculatorElement>();
            ZbDatabase database = new ZbDatabase();
            database.Open(@"C:\Users\Admin\Documents\Новая папка\karta2_uch.zb");
            foreach (var SysSelected in sysHashSet)
            {
                var fields = new[] { "Proklad", "Dpod", "Dobr", "L", "Dw_pod", "Dw_obr", "Texp_nad", "DateExpl" };
                Dictionary<string, string> pairs = new Dictionary<string, string>();
                var obj = database.SelectByKey(SysSelected);
                foreach (var field in fields)
                {
                    int index = obj.GetFieldIndexByName(zbNameType.zbShortName, field);
                    string fieldValue = database.SelectByKey(SysSelected).FieldValue[index];
                    pairs[field] = fieldValue;
                }
                // proklad 1 - надземная, 2 - подземная канальная, 4 - подвальная труба, 11 - подземная бесканальная ПИ-труба

                int dateExpl_year = int.Parse(pairs["DateExpl"].Split(' ')[0].Split('.')[2]) ;
                int texp_nad = int.Parse(pairs["Texp_nad"]);
                int group_num;

                int proklad = int.Parse(pairs["Proklad"]);

                double? dpod; if (double.TryParse(pairs["Dpod"].Replace('.', ','), out var temp)) dpod = temp; else dpod = null;
                double? dobr; if (double.TryParse(pairs["Dobr"].Replace('.', ','), out temp)) dobr = temp; else dobr = null;
                double? dw_pod; if (double.TryParse(pairs["Dw_pod"].Replace('.', ','), out temp)) dw_pod = temp; else dw_pod = null;
                double? dw_obr; if (double.TryParse(pairs["Dw_obr"].Replace('.', ','), out temp)) dw_obr = temp; else dw_obr = null;
                double l = double.Parse(pairs["L"].Replace('.', ','));

                double? f_pod = 0.785 * Math.Pow((double)(dpod - 2f * dw_pod / 1000f), 2);
                double? v_pod = f_pod * l;
                double? f_obr = 0.785 * Math.Pow((double)(dpod - 2f * dw_pod / 1000f), 2);
                double? v_obr = f_obr * l;
                double? me_pod = null;
                double? me_obr = null;

                double? m_pod = null;
                double? P_pod = null;
                double? m_obr = null;
                double? P_obr = null;

                if (proklad == 11 && dateExpl_year >= 1995)
                {
                    group_num = 1;
                    m_pod = m_obr = 0.15;
                    P_pod = P_obr = 0.03;
                }
                else if ((proklad == 2 || proklad == 1 || proklad == 4) && dateExpl_year >= 1997)
                {
                    group_num = 2;
                    m_pod = m_obr = 0.3;
                    P_pod = P_obr = 0.03;
                }
                else if ((proklad == 1 || proklad == 4) && dateExpl_year < 1997)
                {
                    group_num = 3;
                    m_pod = m_obr = 0.3;
                    P_pod = P_obr = 0.07;
                }
                else if (proklad == 2 && dateExpl_year < 1997)
                {
                    group_num = 4;
                    if (dpod is null) ;
                    else if (dpod <= 0.259)
                    {
                        m_pod = 1;
                        P_pod = 0.2;
                        
                    }
                    else if (dpod <= 1.398)
                    {
                        m_pod = 0.85;
                        P_pod = 0.1;
                    }
                    else throw new Exception();

                    if (dobr is null) ;
                    else if (dobr <= 0.259)
                    {
                        m_obr = 1;
                        P_obr = 0.2;
                    }
                    else if (dobr <= 1.398)
                    {
                        m_obr = 0.85;
                        P_obr = 0.1;
                        
                    }
                    else throw new Exception();
                }
                else throw new Exception();

                double? K_pod = null;
                double Vr_pod = 0;
                if (dpod != null)
                {
                    K_pod = 3 * Math.Pow((DateTime.Now.Year - dateExpl_year) / ((double)dw_pod / (double)P_pod), 2.6);
                    if (K_pod > 3) K_pod = 3;
                    me_pod = (1 + K_pod) * m_pod;
                    Vr_pod = (double)(me_pod * v_pod);
                }

                double? K_obr = null;
                double Vr_obr = 0;
                if (dobr != null)
                {
                    K_obr = 3 * Math.Pow((DateTime.Now.Year - dateExpl_year) / ((double)dw_obr / (double)P_obr), 2.6);
                    if (K_obr > 3) K_obr = 3;
                    me_obr = (1 + K_obr) * m_obr;
                    Vr_obr = (double)(me_obr * v_obr);
                }
                double v = Vr_obr + Vr_pod;

                elements.Add( new HeatNetworkWaterVolumeCalculatorElement(dateExpl_year, texp_nad, group_num, proklad, dpod, dobr, dw_pod, dw_obr, l ,f_pod, f_obr, v_pod, v_obr, me_pod, me_obr, m_pod, m_obr, P_pod, P_obr, K_pod, K_obr, v));
            }

            elements.Sort();

            string filePath = "output.xlsx";

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Sheet1");

                var headerCells = worksheet.Range("A1:S3");
                headerCells.Style.Font.Bold = true;
                headerCells.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                headerCells.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                headerCells.Style.Alignment.WrapText = true;
                worksheet.Row(1).Height = 40;

                worksheet.Cell("A1").Value = "Способ прокладки";
                worksheet.Cell("B1").Value = "Год ввода в эксплуатацию";
                worksheet.Cell("C1").Value = "Срок эксплуатации";
                worksheet.Cell("D1").Value = "Номер группы";
                worksheet.Cell("E1").Value = "Внутренний диаметр трубопровода";
                worksheet.Cell("F1").Value = "Толщина стенки трубопровода";
                worksheet.Cell("G1").Value = "Длина трубопровода";
                worksheet.Cell("H1").Value = "";
                worksheet.Cell("I1").Value = "Площадь поперечного сечения";
                worksheet.Cell("J1").Value = "";
                worksheet.Cell("K1").Value = "Объем воды в трубопроводах";
                worksheet.Cell("L1").Value = "";
                worksheet.Cell("M1").Value = "Поправочный коэффициент к фактическому объему";
                worksheet.Cell("N1").Value = "";
                worksheet.Cell("O1").Value = "Расчетный объем воды";

                worksheet.Range("A1:A3").Merge();
                worksheet.Range("B1:B3").Merge();
                worksheet.Range("C1:C3").Merge();
                worksheet.Range("D1:D3").Merge();
                worksheet.Range("E1:E2").Merge();
                worksheet.Range("F1:F2").Merge();
                worksheet.Range("E3:F3").Merge().Value = "под.";

                worksheet.Range("G1:H1").Merge();
                worksheet.Range("G2:H2").Merge().Value = "отопительный";
                worksheet.Cell("G3").Value = "под.";
                worksheet.Cell("H3").Value = "обр.";

                worksheet.Range("I1:J1").Merge();
                worksheet.Range("I2:I3").Merge().Value = "под.";
                worksheet.Range("J2:J3").Merge().Value = "обр.";

                worksheet.Range("K1:L1").Merge();
                worksheet.Range("K2:K3").Merge().Value = "под.";
                worksheet.Range("L2:L3").Merge().Value = "обр.";

                worksheet.Range("M1:N1").Merge();
                worksheet.Range("M2:M3").Merge().Value = "под.";
                worksheet.Range("N2:N3").Merge().Value = "обр.";

                worksheet.Range("O1:P1").Merge();
                worksheet.Range("O2:P3").Merge().Value = "отопительный, V";

                

                int row = 4;

                int stage = -1;

                foreach (var element in elements)
                {
                    int es = element.stage;

                    if (stage != es)
                    {
                        if (stage == -1 && stage != es)
                        {
                            stage++;
                            worksheet.Cell(row++, 1).SetValue("Надземная:").Style.Font.SetBold(true).Font.SetFontSize(12);
                            worksheet.Cell(row++, 1).SetValue("  б) группа II:").Style.Font.SetBold(true).Font.SetFontSize(12);
                        }

                        if (stage == 0 && stage != es)
                        {
                            stage++;
                            worksheet.Cell(row++, 1).SetValue("  в) группа III:").Style.Font.SetBold(true).Font.SetFontSize(12);
                        }

                        if (stage == 1 && stage != es)
                        {
                            stage++;
                            worksheet.Cell(row++, 1).SetValue("Подземная:").Style.Font.SetBold(true).Font.SetFontSize(12);
                            worksheet.Cell(row++, 1).SetValue("  а) группа I:").Style.Font.SetBold(true).Font.SetFontSize(12);
                        }

                        if (stage == 2 && stage != es)
                        {
                            stage++;
                            worksheet.Cell(row++, 1).SetValue("  б) группа II:").Style.Font.SetBold(true).Font.SetFontSize(12);
                        }

                        if (stage == 3 && stage != es)
                        {
                            stage++;
                            worksheet.Cell(row++, 1).SetValue("  г) группа IV:").Style.Font.SetBold(true).Font.SetFontSize(12);
                            worksheet.Cell(row++, 1).SetValue("    1) dвн 0,259 - 1,398").Style.Font.SetBold(true).Font.SetFontSize(12);
                        }

                        if (stage == 4 && stage != es)
                        {
                            stage++;
                            worksheet.Cell(row++, 1).SetValue("    2) dвн < 0,259").Style.Font.SetBold(true).Font.SetFontSize(12);
                        }

                        if (stage == 5 && stage != es)
                        {
                            stage++;
                            worksheet.Cell(row++, 1).SetValue("подвальная труба").Style.Font.SetBold(true).Font.SetFontSize(12);
                            worksheet.Cell(row++, 1).SetValue("  б) группа II").Style.Font.SetBold(true).Font.SetFontSize(12);
                        }

                        if (stage == 6 && stage != es)
                        {
                            stage++;
                            worksheet.Cell(row++, 1).SetValue("  в) группа III").Style.Font.SetBold(true).Font.SetFontSize(12);
                        }
                    }

                    string[] strings = element.GetArray();
                    for (int i = 0; i < strings.Length; i++)
                    {
                        worksheet.Cell(row, i + 1).Value = strings[i];
                    }
                    row++;
                }
                worksheet.Cell(row, 15).FormulaA1 = $"SUM(O4:O{row - 1})";
                worksheet.Cell(row, 15).Style.Font.SetBold(true);
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
