using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO.Ports;
using System.Management;
using Microsoft.Win32;
using System.Security.Principal;
using System.Collections;

namespace SerialLister
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.AutoSize = true;
            this.AutoSizeMode = AutoSizeMode.GrowAndShrink;

            GetSerialPorts();
        }

        public void GetSerialPorts()
        {
            listBox1.Items.Clear();
            var list = new List<KeyValuePair<int, string>>();
            int i = 0;

            using (ManagementClass i_Entity = new ManagementClass("Win32_PnPEntity"))
            {
                foreach (ManagementObject i_Inst in i_Entity.GetInstances())
                { 
                    Object o_Guid = i_Inst.GetPropertyValue("ClassGuid");
                    if (o_Guid == null || o_Guid.ToString().ToUpper() != "{4D36E978-E325-11CE-BFC1-08002BE10318}")
                        continue; // Skip all devices except device class "PORTS"

                    String s_Caption = i_Inst.GetPropertyValue("Caption").ToString();
                    String s_Manufact = i_Inst.GetPropertyValue("Manufacturer").ToString();
                    String s_DeviceID = i_Inst.GetPropertyValue("PnpDeviceID").ToString();
                    String s_RegPath = "HKEY_LOCAL_MACHINE\\System\\CurrentControlSet\\Enum\\" + s_DeviceID + "\\Device Parameters";
                    String s_PortName = Registry.GetValue(s_RegPath, "PortName", "").ToString();

                    int s32_Pos = s_Caption.IndexOf(" (COM");
                    if (s32_Pos > 0) // remove COM port from description
                        s_Caption = s_Caption.Substring(0, s32_Pos);

                    var sp = s_PortName + "\t  " + s_Caption;
                    int spn = Convert.ToInt32(s_PortName.Remove(0, 3));

                    list.Add(new KeyValuePair<int, string>(spn, sp));
                    i++;
                }
            }
            list = list.OrderBy(kvp => kvp.Key).ToList();
            foreach (var element in list)
            {
                listBox1.Items.Add(element.Value);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            GetSerialPorts();
        }
    }
}
