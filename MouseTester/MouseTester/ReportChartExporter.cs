using System;
using System.Collections.Generic;
using System.IO;
using OxyPlot;
using OxyPlot.Annotations;
using OxyPlot.Axes;
using OxyPlot.Series;

namespace MouseTester
{
    /// <summary>
    /// Builds standalone OxyPlot models for the stability report and exports them
    /// as 3840×2160 PNG bytes (Dpi=192, equivalent to 1080p HiDPI content).
    /// </summary>
    internal static class ReportChartExporter
    {
        private const int    ExportWidth  = 3840;
        private const int    ExportHeight = 2160;
        private const float  ExportDpi    = 192f;   // 96 DPI × 200 %
        private const double MarkerSize   = 2.0;
        private const double MmPerInch    = 25.4;

        // ── Public export methods ──────────────────────────────────────────

        /// <summary>Interval vs. Time chart with optional red threshold line.</summary>
        public static byte[] ExportIntervalChart(
            IList<MouseEvent> events, int startIdx, int endIdx,
            string title = null, double? thresholdLine = null)
        {
            var model = CreateBaseModel(title ?? "Interval vs. Time");
            AddAxis(model, AxisPosition.Bottom, "Time (ms)");
            AddAxis(model, AxisPosition.Left,   "Interval (ms)");

            var scatter = CreateScatter(OxyColors.SteelBlue);
            model.Series.Add(scatter);

            for (int i = startIdx + 1; i <= endIdx; i++)
            {
                double t  = events[i].ts;
                double iv = events[i].ts - events[i - 1].ts;
                scatter.Points.Add(new ScatterPoint(t, iv));
            }

            if (thresholdLine.HasValue)
            {
                model.Annotations.Add(new LineAnnotation
                {
                    Type            = LineAnnotationType.Horizontal,
                    Y               = thresholdLine.Value,
                    Color           = OxyColors.Red,
                    StrokeThickness = 3,
                    LineStyle       = LineStyle.Dash,
                    Text            = $"Drop threshold ({thresholdLine.Value:F4} ms)",
                    TextColor       = OxyColors.Red,
                    FontSize        = 18
                });
            }

            return Render(model);
        }

        /// <summary>xCount and yCount vs. Time (blue/red scatter).</summary>
        public static byte[] ExportCountChart(
            IList<MouseEvent> events, int startIdx, int endIdx,
            string title = null)
        {
            var model = CreateBaseModel(title ?? "Count vs. Time");
            AddAxis(model, AxisPosition.Bottom, "Time (ms)");
            AddAxis(model, AxisPosition.Left,   "Counts");

            var blueScatter = CreateScatter(OxyColors.SteelBlue, "xCount");
            var redScatter  = CreateScatter(OxyColors.Crimson,   "yCount");
            model.Series.Add(blueScatter);
            model.Series.Add(redScatter);

            for (int i = startIdx; i <= endIdx; i++)
            {
                double t = events[i].ts;
                blueScatter.Points.Add(new ScatterPoint(t, events[i].lastx));
                redScatter.Points.Add(new ScatterPoint(t,  events[i].lasty));
            }

            model.IsLegendVisible = true;
            model.Legends.Add(new OxyPlot.Legends.Legend
            {
                LegendPosition = OxyPlot.Legends.LegendPosition.TopRight,
                LegendFontSize = 20
            });

            return Render(model);
        }

        /// <summary>Displacement magnitude vs. Time with optional jump threshold line.</summary>
        public static byte[] ExportVelocityChart(
            IList<MouseEvent> events, int startIdx, int endIdx,
            double cpi, string title = null, double? jumpThreshold = null)
        {
            bool useVelocity = cpi > 0;
            string yLabel = useVelocity ? "Velocity (m/s)" : "Displacement (counts)";

            var model = CreateBaseModel(title ?? (useVelocity ? "Velocity vs. Time" : "Displacement vs. Time"));
            AddAxis(model, AxisPosition.Bottom, "Time (ms)");
            AddAxis(model, AxisPosition.Left,   yLabel);

            var scatter = CreateScatter(OxyColors.SteelBlue);
            model.Series.Add(scatter);

            for (int i = startIdx; i <= endIdx; i++)
            {
                double t = events[i].ts;
                double y;
                if (useVelocity && i > startIdx)
                {
                    double dt = events[i].ts - events[i - 1].ts;
                    double disp = Math.Sqrt((double)events[i].lastx * events[i].lastx +
                                            (double)events[i].lasty * events[i].lasty);
                    y = dt > 0 ? disp / dt / cpi * MmPerInch : 0;
                }
                else
                {
                    y = Math.Sqrt((double)events[i].lastx * events[i].lastx +
                                   (double)events[i].lasty * events[i].lasty);
                }
                scatter.Points.Add(new ScatterPoint(t, y));
            }

            if (jumpThreshold.HasValue)
            {
                // For displacement chart, the threshold is in count units
                double displayThreshold = useVelocity ? double.NaN : jumpThreshold.Value;
                if (!double.IsNaN(displayThreshold))
                {
                    model.Annotations.Add(new LineAnnotation
                    {
                        Type            = LineAnnotationType.Horizontal,
                        Y               = displayThreshold,
                        Color           = OxyColors.Red,
                        StrokeThickness = 3,
                        LineStyle       = LineStyle.Dash,
                        Text            = $"Jump threshold ({displayThreshold:F1} counts)",
                        TextColor       = OxyColors.Red,
                        FontSize        = 18
                    });
                }
            }

            return Render(model);
        }

        // ── Internal helpers ───────────────────────────────────────────────

        private static PlotModel CreateBaseModel(string title)
        {
            return new PlotModel
            {
                Title           = title,
                Background      = OxyColors.White,
                TitleFontSize   = 28,
                DefaultFontSize = 20,
                PlotMargins     = new OxyThickness(80, 20, 20, 60)
            };
        }

        private static void AddAxis(PlotModel model, AxisPosition pos, string title)
        {
            model.Axes.Add(new LinearAxis
            {
                Position      = pos,
                Title         = title,
                TitleFontSize = 22,
                FontSize      = 18,
                MajorGridlineStyle = LineStyle.Dot,
                MajorGridlineColor = OxyColor.FromArgb(60, 0, 0, 0)
            });
        }

        private static ScatterSeries CreateScatter(OxyColor color, string title = null)
        {
            return new ScatterSeries
            {
                MarkerFill            = color,
                MarkerSize            = MarkerSize,
                MarkerStroke          = color,
                MarkerStrokeThickness = 0,
                MarkerType            = MarkerType.Circle,
                Title                 = title
            };
        }

        private static byte[] Render(PlotModel model)
        {
            using var ms       = new MemoryStream();
            var exporter = new OxyPlot.SkiaSharp.PngExporter
            {
                Width  = ExportWidth,
                Height = ExportHeight,
                Dpi    = ExportDpi
            };
            exporter.Export(model, ms);
            return ms.ToArray();
        }
    }
}
