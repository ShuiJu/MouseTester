using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using OxyPlot;
using OxyPlot.Axes;
using OxyPlot.Series;
using OxyPlot.WindowsForms;

namespace MouseTester
{
    public partial class MousePlot : Form
    {
        // ── Constants ──────────────────────────────────────────────────────────
        private const double SmoothingHz       = 125.0;
        private const double MmPerInch         = 25.4;
        private const double MarkerSize        = 1.5;
        private const double LineStroke        = 2.0;
        private const double StemStroke        = 1.0;

        // ── Data ───────────────────────────────────────────────────────────────
        private MouseLog mlog;

        // Time range
        private int    last_start;
        private double last_start_time;
        private int    last_end;
        private double last_end_time;

        // Axis bounds computed during RecalculateSeriesData
        private double x_min, x_max, y_min, y_max;
        private string xlabel = "", ylabel = "";

        // ── Persistent series instances (optimization: reuse, don't recreate) ──
        private ScatterSeries _blueScatter, _redScatter;
        // Fit lines (smoothed, shown when Lines checkbox is OFF)
        private LineSeries _blueFitLine, _redFitLine;
        // Raw lines (straight connections, shown when Lines checkbox is ON)
        private LineSeries _blueRawLine, _redRawLine;
        // Stem lines (NaN-break LineSeries, OxyPlot 2.x compatible)
        private LineSeries _blueStem, _redStem;

        private bool _useDualSeries;

        // ── Constructor ────────────────────────────────────────────────────────
        public MousePlot(MouseLog Mlog)
        {
            this.Text = $"MousePlot v{Program.version}";
            InitializeComponent();

            // Deep-copy the log (preserve the original shift-by-one behavior)
            this.mlog = new MouseLog();
            this.mlog.Desc = Mlog.Desc;
            this.mlog.Cpi  = Mlog.Cpi;

            int i         = 1;
            int x         = Mlog.Events[0].lastx;
            int y         = Mlog.Events[0].lasty;
            ushort bflags = Mlog.Events[0].buttonflags;
            double ts     = Mlog.Events[0].ts;
            while (i < Mlog.Events.Count)
            {
                this.mlog.Add(new MouseEvent(bflags, x, y, ts));
                x      = Mlog.Events[i].lastx;
                y      = Mlog.Events[i].lasty;
                bflags = Mlog.Events[i].buttonflags;
                ts     = Mlog.Events[i].ts;
                i++;
            }
            this.mlog.Add(new MouseEvent(bflags, x, y, ts));

            this.last_end      = mlog.Events.Count - 1;
            this.last_end_time = Mlog.Events[this.last_end].ts;
            this.last_start_time = mlog.Events[last_end].ts > 100 ? 100 : 0;

            InitializeSeries();
            InitializePlotModel();

            // Wire up controls (add handlers AFTER setting initial values)
            comboBoxPlotType.SelectedIndex = 0;
            comboBoxPlotType.SelectedIndexChanged += comboBox1_SelectedIndexChanged;

            numericUpDownStart.Minimum      = 0;
            numericUpDownStart.Maximum      = (decimal)last_end_time;
            numericUpDownStart.Value        = (decimal)last_start_time;
            numericUpDownStart.DecimalPlaces = 3;
            numericUpDownStart.Increment    = 10;
            numericUpDownStart.ValueChanged += numericUpDownStart_ValueChanged;

            numericUpDownEnd.Minimum      = 0;
            numericUpDownEnd.Maximum      = (decimal)last_end_time;
            numericUpDownEnd.Value        = (decimal)last_end_time;
            numericUpDownEnd.DecimalPlaces = 3;
            numericUpDownEnd.Increment    = 10;
            numericUpDownEnd.ValueChanged += numericUpDownEnd_ValueChanged;

            checkBoxStem.Checked        = false;
            checkBoxStem.CheckedChanged += checkBoxStem_CheckedChanged;

            checkBoxLines.Checked        = false;
            checkBoxLines.CheckedChanged += checkBoxLines_CheckedChanged;

            // Open to Interval vs. Time
            comboBoxPlotType.SelectedItem = "Interval vs. Time";
            // ^ Triggers comboBox1_SelectedIndexChanged → UpdateYRangeOptions + refresh_plot
        }

        // ── Series initialisation ──────────────────────────────────────────────
        private void InitializeSeries()
        {
            _blueScatter = new ScatterSeries
            {
                MarkerFill            = OxyColors.Blue,
                MarkerSize            = MarkerSize,
                MarkerStroke          = OxyColors.Blue,
                MarkerStrokeThickness = 1.0,
                MarkerType            = MarkerType.Circle
            };
            _redScatter = new ScatterSeries
            {
                MarkerFill            = OxyColors.Red,
                MarkerSize            = MarkerSize,
                MarkerStroke          = OxyColors.Red,
                MarkerStrokeThickness = 1.0,
                MarkerType            = MarkerType.Circle
            };

            _blueFitLine = new LineSeries { Color = OxyColors.Blue, StrokeThickness = LineStroke };
            _redFitLine  = new LineSeries { Color = OxyColors.Red,  StrokeThickness = LineStroke };
            _blueRawLine = new LineSeries { Color = OxyColors.Blue, StrokeThickness = LineStroke };
            _redRawLine  = new LineSeries { Color = OxyColors.Red,  StrokeThickness = LineStroke };
            _blueStem    = new LineSeries { Color = OxyColors.Blue, StrokeThickness = StemStroke };
            _redStem     = new LineSeries { Color = OxyColors.Red,  StrokeThickness = StemStroke };
        }

        private void InitializePlotModel()
        {
            plot1.Model = new PlotModel
            {
                Title      = mlog.Desc,
                PlotType   = PlotType.Cartesian,
                Background = OxyColors.White,
                Subtitle   = mlog.Cpi + " cpi"
            };
        }

        // ── Refresh ────────────────────────────────────────────────────────────
        private void refresh_plot()
        {
            FindStartEndIndices();
            RecalculateSeriesData();
            RebuildPlotWithCurrentOptions();
        }

        private void FindStartEndIndices()
        {
            for (int j = 0; j < mlog.Events.Count; j++)
            {
                if (mlog.Events[j].ts >= last_start_time)
                {
                    last_start      = j;
                    last_start_time = mlog.Events[j].ts;
                    numericUpDownStart.Value = (decimal)last_start_time;
                    break;
                }
            }
            for (int j = mlog.Events.Count - 1; j >= 0; j--)
            {
                if (mlog.Events[j].ts <= last_end_time)
                {
                    last_end      = j;
                    last_end_time = mlog.Events[j].ts;
                    numericUpDownEnd.Value = (decimal)last_end_time;
                    break;
                }
            }
        }

        // ── Data calculation ───────────────────────────────────────────────────
        private void RecalculateSeriesData()
        {
            // Reset all series
            _blueScatter.Points.Clear(); _redScatter.Points.Clear();
            _blueFitLine.Points.Clear(); _redFitLine.Points.Clear();
            _blueRawLine.Points.Clear(); _redRawLine.Points.Clear();
            _blueStem.Points.Clear();    _redStem.Points.Clear();
            _useDualSeries = false;

            // Reset line colours (they may have been changed for single-axis plots)
            _blueFitLine.Color = OxyColors.Blue; _redFitLine.Color = OxyColors.Red;
            _blueRawLine.Color = OxyColors.Blue; _redRawLine.Color = OxyColors.Red;

            reset_minmax();

            string plotType = comboBoxPlotType.Text;
            bool showStats  = plotType.Contains("Interval") || plotType.Contains("Frequency");
            statisticsGroupBox.Visible = showStats;
            groupBoxYRange.Visible     = showStats;

            if (plotType.Contains("xyCount"))
            {
                xlabel = "Time (ms)";
                ylabel = "Counts [x = Blue, y = Red]";
                _useDualSeries = true;
                FillSingleAxisSeries(i => mlog.Events[i].lastx, _blueScatter, _blueRawLine, _blueStem);
                FillSingleAxisSeries(i => mlog.Events[i].lasty, _redScatter,  _redRawLine,  _redStem);
                ApplyFit(_blueScatter, _blueFitLine);
                ApplyFit(_redScatter,  _redFitLine);
            }
            else if (plotType.Contains("xCount"))
            {
                xlabel = "Time (ms)"; ylabel = "xCounts";
                FillSingleAxisSeries(i => mlog.Events[i].lastx, _blueScatter, _blueRawLine, _blueStem);
                ApplyFit(_blueScatter, _blueFitLine);
                _blueFitLine.Color = OxyColors.Green;
            }
            else if (plotType.Contains("yCount"))
            {
                xlabel = "Time (ms)"; ylabel = "yCounts";
                FillSingleAxisSeries(i => mlog.Events[i].lasty, _blueScatter, _blueRawLine, _blueStem);
                ApplyFit(_blueScatter, _blueFitLine);
                _blueFitLine.Color = OxyColors.Green;
            }
            else if (plotType.Contains("Interval") || plotType.Contains("Frequency"))
            {
                FillIntervalSeries(plotType.Contains("Interval"));
                _blueFitLine.Color = OxyColors.Green;
            }
            else if (plotType.Contains("xyVelocity"))
            {
                xlabel = "Time (ms)"; ylabel = "Velocity (m/s) [x = Blue, y = Red]";
                _useDualSeries = true;
                if (mlog.Cpi <= 0) { MessageBox.Show("CPI value is invalid, please run Measure"); return; }
                FillVelocitySeries(e => e.lastx, _blueScatter, _blueRawLine, _blueStem);
                FillVelocitySeries(e => e.lasty, _redScatter,  _redRawLine,  _redStem);
                ApplyFit(_blueScatter, _blueFitLine);
                ApplyFit(_redScatter,  _redFitLine);
            }
            else if (plotType.Contains("xVelocity"))
            {
                xlabel = "Time (ms)"; ylabel = "xVelocity (m/s)";
                if (mlog.Cpi <= 0) { MessageBox.Show("CPI value is invalid, please run Measure"); return; }
                FillVelocitySeries(e => e.lastx, _blueScatter, _blueRawLine, _blueStem);
                ApplyFit(_blueScatter, _blueFitLine);
                _blueFitLine.Color = OxyColors.Green;
            }
            else if (plotType.Contains("yVelocity"))
            {
                xlabel = "Time (ms)"; ylabel = "yVelocity (m/s)";
                if (mlog.Cpi <= 0) { MessageBox.Show("CPI value is invalid, please run Measure"); return; }
                FillVelocitySeries(e => e.lasty, _blueScatter, _blueRawLine, _blueStem);
                ApplyFit(_blueScatter, _blueFitLine);
                _blueFitLine.Color = OxyColors.Green;
            }
            else if (plotType.Contains("X vs. Y"))
            {
                FillXvsYSeries();
            }
        }

        // ── Series fill helpers ────────────────────────────────────────────────
        private void FillSingleAxisSeries(
            Func<int, double> valueAt,
            ScatterSeries scatter, LineSeries rawLine, LineSeries stem)
        {
            for (int i = last_start; i <= last_end; i++)
            {
                double x = mlog.Events[i].ts;
                double y = valueAt(i);
                update_minmax(x, y);
                scatter.Points.Add(new ScatterPoint(x, y));
                rawLine.Points.Add(new DataPoint(x, y));
                AddStemPoint(stem, x, y);
            }
        }

        private void FillVelocitySeries(
            Func<MouseEvent, int> axisSelector,
            ScatterSeries scatter, LineSeries rawLine, LineSeries stem)
        {
            for (int i = last_start; i <= last_end; i++)
            {
                double x = mlog.Events[i].ts;
                double y = i == 0 ? 0.0 : CalcVelocity(i, axisSelector);
                update_minmax(x, y);
                scatter.Points.Add(new ScatterPoint(x, y));
                rawLine.Points.Add(new DataPoint(x, y));
                AddStemPoint(stem, x, y);
            }
        }

        private double CalcVelocity(int i, Func<MouseEvent, int> axisSelector)
        {
            double dt = mlog.Events[i].ts - mlog.Events[i - 1].ts;
            if (dt == 0) return 0;
            return axisSelector(mlog.Events[i]) / dt / mlog.Cpi * MmPerInch;
        }

        private void FillXvsYSeries()
        {
            xlabel = "xCounts"; ylabel = "yCounts";
            double cx = 0, cy = 0;
            for (int i = last_start; i <= last_end; i++)
            {
                cx += mlog.Events[i].lastx;
                cy += mlog.Events[i].lasty;
                update_minmax(cx, cx);
                update_minmax(cy, cy);
                _blueScatter.Points.Add(new ScatterPoint(cx, cy));
                _blueRawLine.Points.Add(new DataPoint(cx, cy));
            }
            _blueRawLine.Color = OxyColors.Green;
        }

        private void FillIntervalSeries(bool isInterval)
        {
            xlabel = "Time (ms)";

            double firstPct, secondPct;
            if (isInterval)
            {
                ylabel = "Update Time (ms)";
                firstPct = 99;   firstPercentileMetricLabel.Text  = "99 Percentile:";
                secondPct = 99.9; secondPercentileMetricLabel.Text = "99.9 Percentile:";
            }
            else
            {
                ylabel = "Frequency (Hz)";
                firstPct = 1;   firstPercentileMetricLabel.Text  = "1 Percentile:";
                secondPct = 0.1; secondPercentileMetricLabel.Text = "0.1 Percentile:";
            }

            Func<double, double> transform = isInterval ? (v => v) : (v => 1000.0 / v);

            var intervals = new List<double>();
            for (int i = last_start; i <= last_end; i++)
            {
                double rawInterval = i == 0 ? 0.0 : (mlog.Events[i].ts - mlog.Events[i - 1].ts);
                double x = mlog.Events[i].ts;
                double y = transform(rawInterval);
                intervals.Add(rawInterval);
                update_minmax(x, y);
                _blueScatter.Points.Add(new ScatterPoint(x, y));
                _blueRawLine.Points.Add(new DataPoint(x, y));
                AddStemPoint(_blueStem, x, y);
            }

            ApplyFit(_blueScatter, _blueFitLine);

            // Statistics
            var desc = intervals.OrderByDescending(v => v).ToList();
            var asc  = intervals.OrderBy(v => v).ToList();
            int count   = intervals.Count;
            int lastIdx = count - 1;
            double avg = transform(desc.Sum() / count);
            double squaredDevs = intervals.Sum(v => Math.Pow(transform(v) - avg, 2));

            double tMax = transform(isInterval ? desc[0]       : desc[lastIdx]);
            double tMin = transform(isInterval ? desc[lastIdx]  : desc[0]);

            maxInterval.Text   = $"{tMax:0.0000####}";
            minInterval.Text   = $"{tMin:0.0000####}";
            avgInterval.Text   = $"{avg:0.0000####}";
            stdevInterval.Text = $"{Math.Sqrt(squaredDevs / lastIdx):0.0000####}";
            rangeInterval.Text = $"{Math.Abs(tMax - tMin):0.0000####}";

            int midIdx = count / 2;
            double median = count % 2 == 1
                ? transform(desc[midIdx])
                : transform((desc[midIdx - 1] + desc[midIdx]) / 2);
            medianInterval.Text = $"{median:0.0000####}";

            var pList = isInterval ? asc : desc;
            firstPercentileInterval.Text  = $"{transform(pList[(int)Math.Ceiling(firstPct  / 100.0 * count) - 1]):0.0000####}";
            secondPercentileInterval.Text = $"{transform(pList[(int)Math.Ceiling(secondPct / 100.0 * count) - 1]):0.0000####}";
        }

        private static void AddStemPoint(LineSeries stemLine, double x, double y)
        {
            stemLine.Points.Add(new DataPoint(x, 0));
            stemLine.Points.Add(new DataPoint(x, y));
            stemLine.Points.Add(new DataPoint(double.NaN, double.NaN)); // break
        }

        // Time-based smoothing fit (125 Hz window)
        private void ApplyFit(ScatterSeries source, LineSeries target)
        {
            target.Points.Clear();
            if (source.Points.Count == 0) return;
            double ms = 1000.0 / SmoothingHz;
            int ind = 0;
            double xMax = source.Points[source.Points.Count - 1].X;
            for (double x = ms; x <= xMax; x += ms)
            {
                double sum = 0;
                int cnt = 0;
                while (ind < source.Points.Count && source.Points[ind].X <= x)
                {
                    sum += source.Points[ind++].Y;
                    cnt++;
                }
                if (cnt > 0)
                    target.Points.Add(new DataPoint(x - ms / 2.0, sum / cnt));
            }
        }

        // ── Plot rebuild (no data recalculation) ───────────────────────────────
        private void RebuildPlotWithCurrentOptions()
        {
            PlotModel pm = plot1.Model;
            pm.Series.Clear();
            pm.Axes.Clear();

            bool isXvsY = comboBoxPlotType.Text.Contains("X vs. Y");

            // Always add scatter
            pm.Series.Add(_blueScatter);
            if (_useDualSeries) pm.Series.Add(_redScatter);

            if (!isXvsY)
            {
                // Line: fit (default) or raw (Lines checkbox)
                if (checkBoxLines.Checked)
                {
                    pm.Series.Add(_blueRawLine);
                    if (_useDualSeries) pm.Series.Add(_redRawLine);
                }
                else
                {
                    pm.Series.Add(_blueFitLine);
                    if (_useDualSeries) pm.Series.Add(_redFitLine);
                }

                // Stem
                if (checkBoxStem.Checked)
                {
                    pm.Series.Add(_blueStem);
                    if (_useDualSeries) pm.Series.Add(_redStem);
                }
            }
            else
            {
                if (checkBoxLines.Checked)
                    pm.Series.Add(_blueRawLine); // green, set in FillXvsYSeries
            }

            // X axis
            double xMargin = x_max > x_min ? (x_max - x_min) / 20.0 : 1.0;
            pm.Axes.Add(new LinearAxis
            {
                Position              = AxisPosition.Bottom,
                Title                 = xlabel,
                AbsoluteMinimum       = x_min - xMargin,
                AbsoluteMaximum       = x_max + xMargin,
                MajorGridlineColor    = OxyColor.FromArgb(40, 0, 0, 139),
                MajorGridlineStyle    = LineStyle.Solid,
                MinorGridlineColor    = OxyColor.FromArgb(20, 0, 0, 139),
                MinorGridlineStyle    = LineStyle.Solid
            });

            // Y axis
            var yAxis = new LinearAxis
            {
                Title              = ylabel,
                MajorGridlineColor = OxyColor.FromArgb(40, 0, 0, 139),
                MajorGridlineStyle = LineStyle.Solid,
                MinorGridlineColor = OxyColor.FromArgb(20, 0, 0, 139),
                MinorGridlineStyle = LineStyle.Solid
            };

            double yPreset = GetYRangePreset();
            if (!double.IsNaN(yPreset))
            {
                yAxis.Minimum        = 0;
                yAxis.Maximum        = yPreset;
                yAxis.AbsoluteMinimum = 0;
                yAxis.AbsoluteMaximum = yPreset;
            }
            else
            {
                double yMargin = y_max > y_min ? (y_max - y_min) / 20.0 : 1.0;
                yAxis.AbsoluteMinimum = y_min - yMargin;
                yAxis.AbsoluteMaximum = y_max + yMargin;
            }
            pm.Axes.Add(yAxis);

            plot1.InvalidatePlot(true);
        }

        // ── Y Range preset ─────────────────────────────────────────────────────
        private double GetYRangePreset()
        {
            if (comboBoxYRange == null || comboBoxYRange.Items.Count == 0) return double.NaN;
            int idx = comboBoxYRange.SelectedIndex;
            if (idx <= 0) return double.NaN;

            if (comboBoxPlotType.Text.Contains("Frequency"))
            {
                double[] presets = { double.NaN, 2000, 4000, 8000, 16000 };
                return idx < presets.Length ? presets[idx] : double.NaN;
            }
            if (comboBoxPlotType.Text.Contains("Interval"))
            {
                double[] presets = { double.NaN, 2.0, 1.0, 0.5, 0.25 };
                return idx < presets.Length ? presets[idx] : double.NaN;
            }
            return double.NaN;
        }

        private void UpdateYRangeOptions()
        {
            comboBoxYRange.SelectedIndexChanged -= comboBoxYRange_SelectedIndexChanged;
            comboBoxYRange.Items.Clear();

            string plotType = comboBoxPlotType.Text;
            if (plotType.Contains("Frequency"))
            {
                comboBoxYRange.Items.AddRange(new object[]
                    { "Auto", "1000 Hz (0~2k)", "2000 Hz (0~4k)", "4000 Hz (0~8k)", "8000 Hz (0~16k)" });
                groupBoxYRange.Visible = true;
            }
            else if (plotType.Contains("Interval"))
            {
                comboBoxYRange.Items.AddRange(new object[]
                    { "Auto", "1ms (0~2ms)", "0.5ms (0~1ms)", "0.25ms (0~0.5ms)", "0.125ms (0~0.25ms)" });
                groupBoxYRange.Visible = true;
            }
            else
            {
                comboBoxYRange.Items.Add("Auto");
                groupBoxYRange.Visible = false;
            }

            comboBoxYRange.SelectedIndex = 0;
            comboBoxYRange.SelectedIndexChanged += comboBoxYRange_SelectedIndexChanged;
        }

        // ── Axis helpers ───────────────────────────────────────────────────────
        private void reset_minmax()
        {
            x_min = double.MaxValue; x_max = double.MinValue;
            y_min = double.MaxValue; y_max = double.MinValue;
        }

        private void update_minmax(double x, double y)
        {
            if (x < x_min) x_min = x;
            if (x > x_max) x_max = x;
            if (y < y_min) y_min = y;
            if (y > y_max) y_max = y;
        }

        // ── Event handlers ─────────────────────────────────────────────────────
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateYRangeOptions();
            refresh_plot();
        }

        private void comboBoxYRange_SelectedIndexChanged(object sender, EventArgs e)
        {
            RebuildPlotWithCurrentOptions();
        }

        private void numericUpDownStart_ValueChanged(object sender, EventArgs e)
        {
            if (numericUpDownStart.Value >= numericUpDownEnd.Value)
                numericUpDownStart.Value = (decimal)last_start_time;
            else
            {
                last_start_time = (double)numericUpDownStart.Value;
                refresh_plot();
            }
        }

        private void numericUpDownEnd_ValueChanged(object sender, EventArgs e)
        {
            if (numericUpDownEnd.Value <= numericUpDownStart.Value)
                numericUpDownEnd.Value = (decimal)last_end_time;
            else
            {
                last_end_time = (double)numericUpDownEnd.Value;
                refresh_plot();
            }
        }

        private void checkBoxStem_CheckedChanged(object sender, EventArgs e)
        {
            RebuildPlotWithCurrentOptions();
        }

        private void checkBoxLines_CheckedChanged(object sender, EventArgs e)
        {
            RebuildPlotWithCurrentOptions();
        }

        // ── PNG export (fixed 3840×2160, white background) ─────────────────────
        // ── Stability Report ───────────────────────────────────────────────────
        private void buttonStability_Click(object sender, EventArgs e)
        {
            if (mlog.Events.Count < 3)
            {
                MessageBox.Show("Not enough data to generate a stability report.",
                    "Stability Report", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // Use the same time window the chart is currently showing
            FindStartEndIndices();

            Cursor = Cursors.WaitCursor;
            try
            {
                AnalysisResult result = MouseAnalysis.Analyze(mlog.Events, last_start, last_end);
                string html = HtmlReportGenerator.Generate(result, mlog.Events, mlog.Desc, mlog.Cpi);

                string reportsDir = Path.Combine(AppContext.BaseDirectory, "reports");
                Directory.CreateDirectory(reportsDir);

                string safeDesc  = string.Concat(mlog.Desc.Split(Path.GetInvalidFileNameChars()));
                string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                string htmlPath  = Path.Combine(reportsDir, $"{safeDesc}_{result.DetectedHz}Hz_{timestamp}.html");

                File.WriteAllText(htmlPath, html, System.Text.Encoding.UTF8);

                // Open in default browser
                Process.Start(new ProcessStartInfo(htmlPath) { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to generate report:\n{ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void buttonSavePNG_Click(object sender, EventArgs e)
        {
            // Default to a "screenshot" folder next to the executable
            string screenshotDir = Path.Combine(AppContext.BaseDirectory, "screenshot");
            Directory.CreateDirectory(screenshotDir);

            // Default filename: {Description}_{3840x2160}_{yyyyMMdd_HHmmss}.png
            string safeDesc = string.Concat(mlog.Desc.Split(Path.GetInvalidFileNameChars()));
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string defaultName = $"{safeDesc}_3840x2160_{timestamp}.png";

            using var dlg = new SaveFileDialog
            {
                Filter           = "PNG Files (*.png)|*.png",
                FilterIndex      = 1,
                InitialDirectory = screenshotDir,
                FileName         = defaultName
            };
            if (dlg.ShowDialog() != DialogResult.OK) return;

            try
            {
                using var stream = File.Create(dlg.FileName);
                var exporter = new OxyPlot.SkiaSharp.PngExporter
                {
                    Width  = 3840,
                    Height = 2160,
                    Dpi    = 192   // 96 DPI × 200%：content scales to match 1080p visual size
                };
                exporter.Export(plot1.Model, stream);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Save failed:\n{ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
