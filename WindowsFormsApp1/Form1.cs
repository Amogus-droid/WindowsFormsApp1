using System.Diagnostics;
using System;
using System.Windows.Forms;
using ZB;
using System.Reflection;
using System.Collections.Generic;
using System.Linq;
using ZuluOcx;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        private DiagnosticArchive diagnosticArchive;
        private HeatNetworkWaterVolumeCalculator heatNetwork;
        private HeatConsumptionSystemWaterVolumeCalculator heatConsumptionSystem;

        public Form1()
        {
            InitializeComponent();
            diagnosticArchive = new DiagnosticArchive(tabControl1.TabPages[0]);
            heatNetwork = new HeatNetworkWaterVolumeCalculator(tabControl1.TabPages[1]);
            heatConsumptionSystem = new HeatConsumptionSystemWaterVolumeCalculator(tabControl1.TabPages[2]);
        }

        private void axMapCtrl1_ObjectSelect(object sender, EventArgs e)
        {
            if (tabControl1.SelectedTab == tabControl1.TabPages[0])
            {
                diagnosticArchive.ObjectSelect(axMapCtrl1.Map.Layers.Active);
            }
            else if (tabControl1.SelectedTab == tabControl1.TabPages[1])
            {
                heatNetwork.ObjectSelect(axMapCtrl1.Map.Layers.Active);
            }
            else if (tabControl1.SelectedTab == tabControl1.TabPages[2])
            {
                heatConsumptionSystem.ObjectSelect(axMapCtrl1.Map.Layers.Active);
            }
        }
    }
}
