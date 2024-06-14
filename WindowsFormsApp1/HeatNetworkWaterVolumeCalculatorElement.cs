using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ZuluLib;

namespace WindowsFormsApp1
{
    internal class HeatNetworkWaterVolumeCalculatorElement : IComparable<HeatNetworkWaterVolumeCalculatorElement>
    {
        public string Proklad { get; }
        public string DateExpl_year { get; }
        public string Texp_nad { get; }
        public string Group_num { get; }
        public string Dpod { get; }
        public string Dw_pod { get; }
        public string L_pod { get; }
        public string L_obr { get; }
        public string F_pod { get; }
        public string F_obr { get; }
        public string V_pod { get; }
        public string V_obr { get; }
        public string Me_pod { get; }
        public string Me_obr { get; }
        public string V { get; }

        public double dpod = 0;
        public double dobr = 0;

        public int stage = 0;

        public HeatNetworkWaterVolumeCalculatorElement(int dateExpl_year, int texp_nad, int group_num, int proklad, double? dpod, double? dobr, double? dw_pod, double? dw_obr, double l, double? f_pod, double? f_obr, double? v_pod, double? v_obr, double? me_pod, double? me_obr, double? pod, double? obr, double? p_pod, double? p_obr, double? k_pod, double? k_obr, double v)
        {
            // proklad 1 - надземная, 2 - подземная канальная, 4 - подвальная труба, 11 - подземная бесканальная ПИ-труба
            switch (proklad)
            {
                case 1:
                    Proklad = "Надземная";
                    break;
                case 2:
                    Proklad = "Подземная канальная";
                    break;
                case 4:
                    Proklad = "Подвальная труба";
                    break;
                case 11:
                    Proklad = "Подземная бесканальная ПИ-труба";
                    break;
            }
            DateExpl_year = dateExpl_year.ToString();
            Texp_nad = (DateTime.Now.Year - dateExpl_year).ToString();
            switch (group_num)
            {
                case 1:
                    Group_num = "I";
                    break;
                case 2:
                    Group_num = "II";
                    break;
                    case 3:
                    Group_num = "III";
                    break;
                case 4:
                    Group_num = "IV";
                    break;
            }            
            Dpod = (dpod * 1000).ToString();
            Dw_pod = dw_pod.ToString();
            if (dpod != null) L_pod = l.ToString(); else L_pod = "0";
            if (dobr != null) L_obr = l.ToString(); else L_obr = "0";

            if (dpod != null) F_pod = f_pod.ToString(); else F_pod = string.Empty;
            if (dobr != null) F_obr = f_obr.ToString(); else F_obr = string.Empty;

            if (dpod != null) V_pod = v_pod.ToString(); else V_pod = string.Empty;
            if (dobr != null) V_obr = v_obr.ToString(); else V_obr = string.Empty;

            if (dpod != null) Me_pod = me_pod.ToString(); else Me_pod = string.Empty;
            if (dobr != null) Me_obr = me_obr.ToString(); else Me_obr = string.Empty;

            if (dpod != null) this.dpod = (double)dpod;
            if (dobr != null) this.dobr = (double)dobr;

            V = v.ToString();

            if (proklad == 1 && group_num == 2) stage = 0;
            else if (proklad == 1 && group_num == 3) stage = 1;
            else if (proklad != 4 && group_num == 1) stage = 2;
            else if (proklad != 4 && group_num == 2) stage = 3;
            else if (proklad != 4 && Math.Max(this.dobr, this.dpod) >= 0.259) stage = 4;
            else if (proklad != 4) stage = 5;
            else if (proklad == 4 && group_num == 2) stage = 6;
            else if (proklad == 4 && group_num == 3) stage = 7;
        }

        public string[] GetArray()
        {
            return new string[]
            {
                Proklad,
                DateExpl_year,
                Texp_nad,
                Group_num,
                Dpod.Replace(",", "."),
                Dw_pod.Replace(",", "."),
                L_pod.Replace(",", "."),
                L_obr.Replace(",", "."),
                F_pod.Replace(",", "."),
                F_obr.Replace(",", "."),
                V_pod.Replace(",", "."),
                V_obr.Replace(",", "."),
                Me_pod.Replace(",", "."),
                Me_obr.Replace(",", "."),
                V.Replace(",", ".")
            };
        }

        public int CompareTo(HeatNetworkWaterVolumeCalculatorElement other)
        {
            return this.stage - other.stage;
        }
    }
}
