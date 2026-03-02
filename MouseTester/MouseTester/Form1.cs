using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;

using OxyPlot;

namespace MouseTester
{
    using OxyPlot.Series;

    public partial class Form1 : Form
    {
        private const double CmPerInch = 2.54;

        private MouseLog mlog = new MouseLog();
        enum state { idle, measure_wait, measure, collect_wait, collect, log };
        private state test_state = state.idle;
        private long pFreq;

        public Form1()
        {
            InitializeComponent();

            this.Text = $"MouseTester v{Program.version}";

            this.RegisterRawInputMouse(Handle);
            this.textBoxDesc.Text = this.mlog.Desc.ToString();
            this.textBoxCPI.Text = this.mlog.Cpi.ToString();
            this.textBox1.Text = "Enter the correct CPI" +
                                 "\r\n        or\r\n" +
                                 "Press the Measure button" +
                                 "\r\n        or\r\n" +
                                 "Press the Load button";
            this.toolStripStatusLabel1.Text = "";
            this.KeyPreview = true;
            this.KeyDown += new KeyEventHandler(Form1_KeyDown);

            QueryPerformanceFrequency(out pFreq);
        }

        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F1)
            {
                buttonLog.PerformClick();
                e.Handled = true;
            }
            if (e.KeyCode == Keys.F2)
            {
                buttonLog.PerformClick();
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.F3)
            {
                buttonPlot.PerformClick();
                e.Handled = false;
            }
        }

        protected override void WndProc(ref Message m)
        {
            QueryPerformanceCounter(out long pCounter);
            if (m.Msg == WM_INPUT)
            {
                RAWINPUT raw = new RAWINPUT();
                uint size = (uint)Marshal.SizeOf(typeof(RAWINPUT));
                int outsize = GetRawInputData(m.LParam, RID_INPUT, out raw, ref size, (uint)Marshal.SizeOf(typeof(RAWINPUTHEADER)));
                if (outsize != -1)
                {
                    if (raw.header.dwType == RIM_TYPEMOUSE)
                    {
                        logMouseEvent(new MouseEvent(raw.data.mouse.buttonsStr.usButtonFlags, raw.data.mouse.lLastX, -(raw.data.mouse.lLastY), pCounter));
                    }
                }
            }
            base.WndProc(ref m);
        }

        private void logMouseEvent(MouseEvent mevent)
        {
            if (this.test_state == state.idle)
            {
            }
            else if (this.test_state == state.measure_wait)
            {
                if (mevent.buttonflags == RI_MOUSE_LEFT_BUTTON_DOWN)
                {
                    this.mlog.Add(mevent);
                    this.toolStripStatusLabel1.Text = "Measuring";
                    this.test_state = state.measure;
                }
            }
            else if (this.test_state == state.measure)
            {
                this.mlog.Add(mevent);
                if (mevent.buttonflags == RI_MOUSE_LEFT_BUTTON_UP)
                {
                    double x = 0.0;
                    double y = 0.0;
                    foreach (MouseEvent e in this.mlog.Events)
                    {
                        x += (double)e.lastx;
                        y += (double)e.lasty;
                    }
                    tsCalc();
                    this.mlog.Cpi = Math.Round(Math.Sqrt((x * x) + (y * y)) / (10 / CmPerInch));
                    this.textBoxCPI.Text = this.mlog.Cpi.ToString();
                    this.textBox1.Text = "Press the Collect or Log Start button\r\n";
                    this.toolStripStatusLabel1.Text = "";
                    this.test_state = state.idle;
                }
            }
            else if (this.test_state == state.collect_wait)
            {
                if (mevent.buttonflags == RI_MOUSE_LEFT_BUTTON_DOWN)
                {
                    this.mlog.Add(mevent);
                    this.toolStripStatusLabel1.Text = "Collecting";
                    this.test_state = state.collect;
                }
            }
            else if (this.test_state == state.collect)
            {
                this.mlog.Add(mevent);
                if (mevent.buttonflags == RI_MOUSE_LEFT_BUTTON_UP)
                {
                    tsCalc();
                    this.textBox1.Text = BuildSummaryText();
                    this.toolStripStatusLabel1.Text = "";
                    this.test_state = state.idle;
                }
            }
            else if (this.test_state == state.log)
            {
                this.mlog.Add(mevent);
            }
        }

        private string BuildSummaryText()
        {
            return "Press the plot button to view data\r\n" +
                   "        or\r\n" +
                   "Press the save button to save log file\r\n" +
                   "Events: " + this.mlog.Events.Count.ToString() + "\r\n" +
                   "Sum X: " + this.mlog.deltaX().ToString() + " counts    " + Math.Abs(this.mlog.deltaX() / this.mlog.Cpi * CmPerInch).ToString("0.0") + " cm\r\n" +
                   "Sum Y: " + this.mlog.deltaY().ToString() + " counts    " + Math.Abs(this.mlog.deltaY() / this.mlog.Cpi * CmPerInch).ToString("0.0") + " cm\r\n" +
                   "Path: " + this.mlog.path().ToString("0") + " counts    " + (this.mlog.path() / this.mlog.Cpi * CmPerInch).ToString("0.0") + " cm";
        }

        private void buttonMeasure_Click(object sender, EventArgs e)
        {
            if (this.test_state == state.idle)
            {
                this.textBox1.Text = "1. Press and hold the left mouse button\r\n" +
                                     "2. Move the mouse 10 cm in a straight line\r\n" +
                                     "3. Release the left mouse button\r\n";
                this.toolStripStatusLabel1.Text = "Press the left mouse button";
                this.mlog.Clear();
                this.test_state = state.measure_wait;
            }
        }

        private void buttonCollect_Click(object sender, EventArgs e)
        {
            if (this.test_state == state.idle)
            {
                this.textBox1.Text = "1. Press and hold the left mouse button\r\n" +
                                     "2. Move the mouse\r\n" +
                                     "3. Release the left mouse button\r\n";
                this.toolStripStatusLabel1.Text = "Press the left mouse button";
                this.mlog.Clear();
                this.test_state = state.collect_wait;
            }
        }

        private void buttonLog_Click(object sender, EventArgs e)
        {
            if (this.test_state == state.idle)
            {
                this.textBox1.Text = "1. Press the Log Stop button\r\n";
                this.toolStripStatusLabel1.Text = "Logging...";
                this.mlog.Clear();
                this.test_state = state.log;
                buttonLog.Text = "Stop (F2)";
            }
            else if (this.test_state == state.log)
            {
                tsCalc();
                this.textBox1.Text = BuildSummaryText();
                this.toolStripStatusLabel1.Text = "";
                this.test_state = state.idle;
                buttonLog.Text = "Start (F1)";
            }
        }

        private void buttonPlot_Click(object sender, EventArgs e)
        {
            if (this.mlog.Events.Count > 0 && this.test_state != state.log && this.test_state != state.collect_wait)
            {
                this.mlog.Desc = textBoxDesc.Text;
                MousePlot mousePlot = new MousePlot(this.mlog);
                mousePlot.Show();
            }
        }

        private void buttonLoad_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "CSV Files (*.csv)|*.csv|All Files(*.*)|*.*";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.Multiselect = false;
            if (openFileDialog1.ShowDialog() != System.Windows.Forms.DialogResult.OK)
                return;

            try
            {
                this.mlog.Load(openFileDialog1.FileName);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to load file:\n{ex.Message}", "Load Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            this.textBox1.Text = BuildSummaryText();
            this.textBoxDesc.Text = this.mlog.Desc.ToString();
            this.textBoxCPI.Text = this.mlog.Cpi.ToString();
            if (this.mlog.Events.Count > 0)
            {
                MousePlot mousePlot = new MousePlot(this.mlog);
                mousePlot.Show();
            }
        }

        private void buttonSave_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "CSV Files (*.csv)|*.csv|All Files(*.*)|*.*";
            saveFileDialog1.FilterIndex = 1;
            if (saveFileDialog1.ShowDialog() != System.Windows.Forms.DialogResult.OK)
                return;

            try
            {
                this.mlog.Desc = textBoxDesc.Text;
                this.mlog.Save(saveFileDialog1.FileName);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to save file:\n{ex.Message}", "Save Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void textBoxCPI_Validated(object sender, EventArgs e)
        {
            try
            {
                this.mlog.Cpi = double.Parse(this.textBoxCPI.Text);
            }
            catch
            {
                MessageBox.Show("Invalid CPI, resetting to previous value");
                this.textBoxCPI.Text = this.mlog.Cpi.ToString();
            }
            this.textBox1.Text = "Press the Collect or Log Start button\r\n";
        }

        private void tsCalc()
        {
            long pcounter_min = this.mlog.Events[0].pcounter;
            foreach (MouseEvent me in this.mlog.Events)
            {
                me.ts = (me.pcounter - pcounter_min) * 1000.0 / pFreq;
            }
        }
    }
}
