using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Windows.Forms;

namespace ComputerInfo
{
    public partial class Form1 : Form
    {
        //保存计算机信息字典，可以随需要保存的信息扩大
        Dictionary<string, string> network = new Dictionary<string, string>
            {
                { "COM_NAME", SystemInformation.ComputerName },
                { "IP", GetIp() },
                { "MAC", GetMac() },
                { "OS", GetOsVer() },
                { "DEPART", ""},
                { "NAME",""},
                { "LOCATION",""},
                { "SERIAL", ""}
            };
        public Form1()
        {
            InitializeComponent();
            GetAll();
        }

        //根据socket连通状态获取本机IP地址
        private static string GetIp()
        {
            try
            {
                Socket s = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
                s.Connect("www.baidu.com", 80);
                IPEndPoint localIpEndPoint = s.LocalEndPoint as IPEndPoint;
                return localIpEndPoint.Address.ToString();
            }
            catch (Exception)
            {
                return null;
            }
        }
        //根据GetIp获取到的IP地址获取相应的MAC地址
        private static string GetMac()
        {
            NetworkInterface[] networkInterface = NetworkInterface.GetAllNetworkInterfaces();
            int i = 0;
            string ip = null, mac = null;
            foreach (NetworkInterface adapter in networkInterface)
            {
                IPInterfaceProperties adapterProperties = adapter.GetIPProperties();
                UnicastIPAddressInformationCollection allAddress = adapterProperties.UnicastAddresses;
                if (allAddress.Count > 0)
                {
                    i++;
                    foreach (UnicastIPAddressInformation addr in allAddress)
                    {
                        if (addr.Address.AddressFamily == AddressFamily.InterNetwork)
                        {
                            ip = addr.Address.ToString();
                        }
                    }
                }

                if (GetIp() == ip)
                {
                    PhysicalAddress address = adapter.GetPhysicalAddress();
                    byte[] bytes = address.GetAddressBytes();
                    for (int j = 0; j < bytes.Length; j++)
                    {
                        mac += (bytes[j].ToString("X2"));
                        if (j != bytes.Length - 1)
                        {
                            mac += '-';
                        }
                    }
                }
            }
            return mac;
        }
        //从注册表获取操作系统版本，此部分后续完善
        private static string GetOsVer()
        {
            return Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProductName", "").ToString() + Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion", "CurrentBuild", "").ToString();
        }

        //获取所有字典需要的信息
        private void GetAll()
        {
            textBox1.Text = network["COM_NAME"];
            textBox2.Text = network["IP"];
            textBox3.Text = network["MAC"];
            textBox4.Text = network["OS"];

        }

        //依次读取所选文件夹内的所有文件到二维数组，判定文件格式。
        //将文件夹里的所有txt文件读入一个list，从list里依次读取文件内容，并将文件名作为字典的一个字段，其余内容依次读入字典
        //当文件夹内的文件全部读取完毕后，将字典写入Excel。

        string[] file_name;
        //存储文件夹内所有符合条件的文件名
        private string[] ReadFileName()
        {
            using (var fbd = new FolderBrowserDialog())
            {
                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    return file_name = Directory.GetFiles(fbd.SelectedPath);
                }
                else
                {
                    MessageBox.Show("请重新选择正确的文件夹。", "Message");
                    return null;
                }
            }
        }
        private void ReadInfo()
        {
            foreach (string currectFile in file_name)
            {
                StreamReader file = new StreamReader(currectFile);
                while ((file.ReadLine()) != null)
                {
                    network["COM_NAME"] = null;
                    network["IP"] = null;
                    network["MAC"] = null;
                    network["OS"] = null;
                    network["DEPART"] = null;
                    network["NAME"] = currectFile;
                    network["LOCATION"] = null;
                    network["SERIAL"] = null;
                }
                file.Close();
            }

        }
        //按钮事件，保存所有窗口上显示的和用户填入的信息
        private void button1_Click(object sender, EventArgs e)
        {
            string departText = textBox5.Text.Trim(),
                    nameText = textBox6.Text.Trim(),
                    locationText = textBox7.Text.Trim(),
                    serialText = textBox8.Text.Trim();
            Boolean a = departText != String.Empty && departText.Length > 1,
                b = nameText != String.Empty && nameText.Length > 1,
                c = locationText != String.Empty && locationText.Length > 1,
                d = serialText != String.Empty && serialText.Length > 1;
            if (a && b && c && d)
            {
                network["DEPART"] = textBox5.Text;
                network["NAME"] = textBox6.Text;
                network["LOCATION"] = textBox7.Text;
                network["SERIAL"] = textBox8.Text;
                using (System.IO.StreamWriter file = new System.IO.StreamWriter(textBox6.Text + ".txt"))
                    foreach (var entry in network)
                        file.WriteLine("{0}: {1}", entry.Key, entry.Value);
                MessageBox.Show("已完成。", "恭喜！");
            }
            else
            {
                MessageBox.Show("请在右侧每个输入框里填入你的相关信息！", "警告！！！");

            }
        }
        //功能待开发
        private void button2_Click(object sender, EventArgs e)
        {
            ReadInfo();
            //this.Close();
        }
    }
}
