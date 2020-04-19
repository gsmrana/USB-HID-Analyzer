using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Globalization;
using HidLibrary;
using System.Diagnostics;
using System.Reflection;

namespace USB_HID_Analyzer
{
    public partial class MainForm : Form
    {
        #region Data

        const int WriteReportTimeout = 3000;

        HidDevice _hidDevice;
        List<HidDevice> _hidDeviceList;

        #endregion

        #region ctor

        public MainForm()
        {
            InitializeComponent();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            try
            {
                rtbEventLog.Font = new Font("Consolas", 9);
            }
            catch (Exception ex)
            {
                PopupException(ex.Message);
            }
        }

        private void MainForm_Shown(object sender, EventArgs e)
        {
            try
            {
                _hidDeviceList = HidDevices.Enumerate().ToList();
                UpdateHidDeviceList();
                toolStripStatusLabel1.Text = "Please select device and click connect to start.";
            }
            catch (Exception ex)
            {
                PopupException(ex.Message);
            }
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {

            }
            catch (Exception ex)
            {
                PopupException(ex.Message);
            }
        }

        #endregion

        #region Internal Methods

        private void AppendEventLog(string str, Color? color = null, bool appendNewLine = true)
        {
            var clr = color ?? Color.Black;
            if (appendNewLine) str += Environment.NewLine;

            // update from UI thread
            Invoke(new MethodInvoker(() =>
            {
                rtbEventLog.SelectionStart = rtbEventLog.TextLength;
                rtbEventLog.SelectionLength = 0;
                rtbEventLog.SelectionColor = clr;
                rtbEventLog.AppendText(str);
                if (!rtbEventLog.Focused) rtbEventLog.ScrollToCaret();
            }));
        }

        private static string ByteArrayToHexString(byte[] bytes, string separator = "")
        {
            return BitConverter.ToString(bytes).Replace("-", separator);
        }

        private static byte[] HexStringToByteArray(string hexstr)
        {
            hexstr.Trim();
            hexstr = hexstr.Replace("-", "");
            hexstr = hexstr.Replace(" ", "");
            return Enumerable.Range(0, hexstr.Length)
                             .Where(x => x % 2 == 0)
                             .Select(x => Convert.ToByte(hexstr.Substring(x, 2), 16))
                             .ToArray();
        }

        private static string GetNTString(byte[] bytes)
        {
            var buffer = Encoding.Unicode.GetString(bytes).ToArray();
            int index = Array.IndexOf(buffer, '\0');
            return new string(buffer, 0, index >= 0 ? index : buffer.Length);
        }

        private void UpdateHidDeviceList()
        {
            var count = 0;
            dataGridView1.SelectionChanged -= DataGridView1_SelectionChanged;
            dataGridView1.Rows.Clear();
            foreach (var d in _hidDeviceList)
            {
                count++;
                var devname = "";
                var devmfg = "";
                var devserial = "";
                var info = d.Attributes;
                var caps = d.Capabilities;
                if (d.ReadProduct(out byte[] bytes)) devname = GetNTString(bytes);
                if (d.ReadManufacturer(out bytes)) devmfg = GetNTString(bytes);
                if (d.ReadSerialNumber(out bytes)) devserial = GetNTString(bytes);

                var row = new string[]
                {
                    count.ToString(),
                    string.Format("VID_{0:X4} PID_{1:X4} REV_{2:X4}", info.VendorId,info.ProductId, info.Version),
                    devname,
                    devmfg,
                    devserial,
                    caps.InputReportByteLength.ToString(),
                    caps.OutputReportByteLength.ToString(),
                    caps.FeatureReportByteLength.ToString(),
                    string.Format("{0:X2}", caps.Usage),
                    string.Format("{0:X4}", caps.UsagePage),
                    d.DevicePath
                };
                dataGridView1.Rows.Add(row);
            }
            dataGridView1.SelectionChanged += DataGridView1_SelectionChanged;
            DataGridView1_SelectionChanged(this, null);
        }

        private void PopupException(string message, string caption = "Exception")
        {
            Invoke(new Action(() =>
            {
                MessageBox.Show(message, caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }));
        }

        #endregion

        #region MenuStrip Events

        private void NewToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(Assembly.GetExecutingAssembly().Location);
            }
            catch (Exception ex)
            {
                PopupException(ex.Message);
            }
        }

        private void ExitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                this.Close();
            }
            catch (Exception ex)
            {
                PopupException(ex.Message);
            }
        }

        private void AboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                HelpToolStripButton_Click(sender, e);
            }
            catch (Exception ex)
            {
                PopupException(ex.Message);
            }
        }

        #endregion

        #region ToolStrip Events

        private void ToolStripButtonReload_Click(object sender, EventArgs e)
        {
            try
            {
                _hidDeviceList = HidDevices.Enumerate().ToList();
                UpdateHidDeviceList();
            }
            catch (Exception ex)
            {
                PopupException(ex.Message);
            }
        }

        private void ToolStripButtonFilter_Click(object sender, EventArgs e)
        {
            try
            {
                var str = toolStripTextBoxVidPid.Text.Split(':');
                var vid = int.Parse(str[0], NumberStyles.AllowHexSpecifier);
                var pid = int.Parse(str[1], NumberStyles.AllowHexSpecifier);
                _hidDeviceList = HidDevices.Enumerate(vid, pid).ToList();
                UpdateHidDeviceList();
            }
            catch (Exception ex)
            {
                PopupException(ex.Message);
            }
        }

        private void ToolStripButtonConnect_Click(object sender, EventArgs e)
        {
            try
            {
                _hidDevice = _hidDeviceList[dataGridView1.SelectedRows[0].Index];
                if (_hidDevice == null) throw new Exception("Could not find Hid USB Device with specified VID PID");

                //open as read write in parallel
                _hidDevice.OpenDevice(DeviceMode.Overlapped, DeviceMode.Overlapped, ShareMode.ShareRead | ShareMode.ShareWrite);
                _hidDevice.Inserted += HidDevice_Inserted;
                _hidDevice.Removed += HidDevice_Removed;
                _hidDevice.MonitorDeviceEvents = true;
            }
            catch (Exception ex)
            {
                PopupException(ex.Message);
            }
        }

        private void ToolStripButtonClear_Click(object sender, EventArgs e)
        {
            try
            {
                rtbEventLog.Clear();
            }
            catch (Exception ex)
            {
                PopupException(ex.Message);
            }
        }

        private void HelpToolStripButton_Click(object sender, EventArgs e)
        {
            try
            {
                var aboutbox = new AboutBox();
                aboutbox.ShowDialog();
            }
            catch (Exception ex)
            {
                PopupException(ex.Message);
            }
        }

        #endregion

        #region Form Events

        private void ButtonReadInput_Click(object sender, EventArgs e)
        {
            try
            {
                var len = _hidDevice.Capabilities.InputReportByteLength;
                if (len <= 0) throw new Exception("This device has no Input Report support!");

                _hidDevice.ReadReport(Device_InputReportReceived);
                AppendEventLog("Hid Input Report Callback Started.");

                //var report = _hidDevice.ReadReport(); //blocking
                //var str = string.Format("Read Input Report [{0}] <-- {1}", report.Data.Length, ByteArrayToHexString(report.Data));
                //AppendLog(str);
            }
            catch (Exception ex)
            {
                PopupException(ex.Message);
            }
        }

        private void ButtonWriteOutput_Click(object sender, EventArgs e)
        {
            try
            {
                var hidReportId = byte.Parse(comboBoxReportId.Text);
                var buf = HexStringToByteArray(textBoxWriteData.Text);

                var report = _hidDevice.CreateReport();
                if (buf.Length > report.Data.Length)
                    throw new Exception("Output Report Length Exceed");

                report.ReportId = hidReportId;
                Array.Copy(buf, report.Data, buf.Length);
                _hidDevice.WriteReport(report, WriteReportTimeout);
                var str = string.Format("Tx Output Report [{0}] --> ID:{1}, {2}", report.Data.Length + 1, report.ReportId, ByteArrayToHexString(report.Data));
                AppendEventLog(str, Color.DarkGreen);
            }
            catch (Exception ex)
            {
                PopupException(ex.Message);
            }
        }

        private void ButtonReadFeature_Click(object sender, EventArgs e)
        {
            try
            {
                var hidReportId = byte.Parse(comboBoxReportId.Text);
                var len = _hidDevice.Capabilities.FeatureReportByteLength;
                if (len <= 0) throw new Exception("This device has no Feature Report support!");

                _hidDevice.ReadFeatureData(out byte[] buf, hidReportId);
                var str = string.Format("Rx Feature Report [{0}] <-- {1}", buf.Length, ByteArrayToHexString(buf));
                AppendEventLog(str, Color.Blue);
            }
            catch (Exception ex)
            {
                PopupException(ex.Message);
            }
        }

        private void ButtonWriteFeature_Click(object sender, EventArgs e)
        {
            try
            {
                var hidReportId = byte.Parse(comboBoxReportId.Text);
                var buf = HexStringToByteArray(textBoxWriteData.Text);

                var len = _hidDevice.Capabilities.FeatureReportByteLength;
                if (buf.Length > len)
                    throw new Exception("Write Feature Report Length Exceed");

                Array.Resize(ref buf, len);
                _hidDevice.WriteFeatureData(buf);
                var str = string.Format("Tx Feature Report [{0}] --> {1}", buf.Length, ByteArrayToHexString(buf));
                AppendEventLog(str, Color.DarkGreen);
            }
            catch (Exception ex)
            {
                PopupException(ex.Message);
            }
        }

        private void DataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView1.SelectedRows.Count <= 0) return;
                var index = dataGridView1.SelectedRows[0].Index;

                var devinfo = _hidDeviceList[index].Attributes;
                toolStripTextBoxVidPid.Text = string.Format("{0:X4}:{1:X4}", devinfo.VendorId, devinfo.ProductId);
            }
            catch (Exception ex)
            {
                PopupException(ex.Message);
            }
        }

        #endregion

        #region Device Event

        private void HidDevice_Inserted()
        {
            try
            {
                var str = string.Format("Inserted  Hid  Device --> VID {0:X4}, PID {0:X4}",
                    _hidDevice.Attributes.VendorId, _hidDevice.Attributes.ProductId);
                AppendEventLog(str);
                //_hidDevice.ReadReport(_device_InputReportReceived);
            }
            catch (Exception ex)
            {
                PopupException(ex.Message);
            }
        }

        private void HidDevice_Removed()
        {
            try
            {
                var str = string.Format("Removed Hid Device --> VID {0:X4}, PID {0:X4}",
                    _hidDevice.Attributes.VendorId, _hidDevice.Attributes.ProductId);
                AppendEventLog(str);
            }
            catch (Exception ex)
            {
                PopupException(ex.Message);
            }
        }

        private void Device_InputReportReceived(HidReport report)
        {
            try
            {
                if (!_hidDevice.IsConnected || !_hidDevice.IsOpen) return;
                var str = string.Format("Rx  Input  Report [{0}] <-- ID:{1}, {2}",
                    report.Data.Length + 1, report.ReportId, ByteArrayToHexString(report.Data));
                AppendEventLog(str, Color.Blue);
                _hidDevice.ReadReport(Device_InputReportReceived);
            }
            catch (Exception ex)
            {
                PopupException(ex.Message);
            }
        }


        #endregion


    }
}
