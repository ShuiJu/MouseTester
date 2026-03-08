using System;
using System.Collections.Generic;
using System.Text;

namespace MouseTester
{
    /// <summary>
    /// Assembles the stability report HTML from an AnalysisResult.
    /// PNG images are Base64-encoded and embedded as data URIs.
    /// </summary>
    internal static class HtmlReportGenerator
    {
        // Chart images are rendered at 3840×2160 with Dpi=192 (= HiDPI 2×).
        // max-width:1920px makes them appear at 1920-pixel logical size in the browser.
        private const string ImgStyle = "style=\"width:100%;max-width:1920px;display:block;margin:12px 0;border:1px solid #ddd;\"";

        public static string Generate(
            AnalysisResult r,
            IList<MouseEvent> events,
            string deviceDesc,
            double cpi)
        {
            var sb = new StringBuilder(512 * 1024);

            // ── Charts ─────────────────────────────────────────────────────
            // Summary: full analysis-range Interval chart
            byte[] summaryPng = ReportChartExporter.ExportIntervalChart(
                events, r.StartIdx, r.EndIdx,
                title: $"Interval vs. Time — {deviceDesc} ({r.DetectedHz} Hz detected)");

            // Drop: ±100 ms window around worst drop (or full range if none)
            byte[] dropPng;
            if (r.DropCount > 0)
            {
                var (ds, de) = MouseAnalysis.GetWindowIndices(events, r.WorstDropIdx, 100, r.StartIdx, r.EndIdx);
                dropPng = ReportChartExporter.ExportIntervalChart(
                    events, ds, de,
                    title: $"Worst Frame Drop ±100 ms  (threshold = {r.TNominal * 1.5:F4} ms)",
                    thresholdLine: r.TNominal * 1.5);
            }
            else
            {
                dropPng = ReportChartExporter.ExportIntervalChart(
                    events, r.StartIdx, r.EndIdx,
                    title: "Interval vs. Time — No Drops Detected",
                    thresholdLine: r.TNominal * 1.5);
            }

            // Null packets: ±100 ms around densest null cluster (or full range)
            byte[] nullPng;
            if (r.NullTotal > 0)
            {
                var (ns, ne) = MouseAnalysis.GetWindowIndices(events, r.WorstNullCenterIdx, 100, r.StartIdx, r.EndIdx);
                nullPng = ReportChartExporter.ExportCountChart(
                    events, ns, ne,
                    title: "xCount / yCount — Null Packet Cluster ±100 ms");
            }
            else
            {
                nullPng = ReportChartExporter.ExportCountChart(
                    events, r.StartIdx, r.EndIdx,
                    title: "xCount / yCount vs. Time — No Null Packets");
            }

            // Fake polling: first 500 ms of analysis range to reveal alternating pattern
            int fakeEnd = r.StartIdx;
            double target = events[r.StartIdx].ts + 500.0;
            while (fakeEnd < r.EndIdx && events[fakeEnd + 1].ts <= target) fakeEnd++;
            byte[] fakePng = ReportChartExporter.ExportIntervalChart(
                events, r.StartIdx, fakeEnd,
                title: "Interval vs. Time — First 500 ms (Fake Polling Pattern Check)");

            // Jump: ±100 ms around worst jump (only if jumps detected)
            byte[] jumpPng = null;
            if (r.JumpCount > 0)
            {
                var (js, je) = MouseAnalysis.GetWindowIndices(events, r.WorstJumpIdx, 100, r.StartIdx, r.EndIdx);
                jumpPng = ReportChartExporter.ExportVelocityChart(
                    events, js, je, cpi,
                    title: "Displacement — Worst Jump Event ±100 ms",
                    jumpThreshold: cpi > 0 ? (double?)null : r.JumpThreshold);
            }

            // ── HTML assembly ──────────────────────────────────────────────
            sb.Append(HtmlHeader(deviceDesc, r));

            // Summary
            sb.Append("<section><h2>Summary</h2>");
            sb.Append(SummaryTable(r));
            sb.Append($"<img src=\"data:image/png;base64,{ToBase64(summaryPng)}\" {ImgStyle}>");
            sb.Append("</section>");

            // Frame Drops
            sb.Append("<section><h2>Frame Drops</h2>");
            if (r.DropCount > 0)
            {
                sb.Append($"<p class=\"warn\">⚠ {r.DropCount} drop event(s) detected " +
                           $"({r.DropRate:F3}%), {r.TotalMissedPolls} poll(s) missed, " +
                           $"max {r.MaxConsecutiveMiss} consecutive.</p>");
            }
            else
            {
                sb.Append("<p class=\"ok\">✓ No frame drops detected.</p>");
            }
            sb.Append($"<img src=\"data:image/png;base64,{ToBase64(dropPng)}\" {ImgStyle}>");
            sb.Append("</section>");

            // Null Packets
            sb.Append("<section><h2>Null Packets</h2>");
            if (r.NullTotal > 0)
            {
                sb.Append($"<p class=\"warn\">⚠ {r.NullTotal} null packet(s) ({r.NullRate:F2}%) — " +
                           $"{r.NullDuringMotion} occurred during movement.</p>");
            }
            else
            {
                sb.Append("<p class=\"ok\">✓ No null packets detected.</p>");
            }
            sb.Append($"<img src=\"data:image/png;base64,{ToBase64(nullPng)}\" {ImgStyle}>");
            sb.Append("</section>");

            // Fake Polling
            sb.Append("<section><h2>Fake Polling Rate</h2>");
            bool suspectFake = r.FakeRatio >= 1.8 || r.DupRate > 30 || r.AltNullRate > 30;
            if (suspectFake)
            {
                sb.Append($"<p class=\"warn\">⚠ Possible fake polling detected: " +
                           $"declared {r.DetectedHz} Hz, effective ≈ {r.EffectiveHz:F0} Hz " +
                           $"(ratio {r.FakeRatio:F2}×). Dup rate {r.DupRate:F1}%, " +
                           $"alternating null {r.AltNullRate:F1}%.</p>");
            }
            else
            {
                sb.Append($"<p class=\"ok\">✓ No fake polling detected. Effective Hz ≈ {r.EffectiveHz:F0} Hz " +
                           $"(ratio {r.FakeRatio:F2}×).</p>");
            }
            sb.Append($"<img src=\"data:image/png;base64,{ToBase64(fakePng)}\" {ImgStyle}>");
            sb.Append("</section>");

            // Jump Events
            sb.Append("<section><h2>Jump Events</h2>");
            if (r.JumpCount > 0 && jumpPng != null)
            {
                sb.Append($"<p class=\"warn\">⚠ {r.JumpCount} jump event(s) detected " +
                           $"(threshold: {r.JumpThreshold:F1} counts).</p>");
                sb.Append($"<img src=\"data:image/png;base64,{ToBase64(jumpPng)}\" {ImgStyle}>");
            }
            else
            {
                sb.Append("<p class=\"ok\">✓ No jump events detected.</p>");
            }
            sb.Append("</section>");

            // Display Compatibility
            sb.Append("<section><h2>Display Compatibility</h2>");
            sb.Append(CompatTable(r));
            sb.Append("</section>");

            sb.Append("</main></body></html>");
            return sb.ToString();
        }

        // ── HTML building blocks ───────────────────────────────────────────

        private static string HtmlHeader(string desc, AnalysisResult r)
        {
            return $@"<!DOCTYPE html>
<html lang=""en"">
<head>
<meta charset=""UTF-8"">
<meta name=""viewport"" content=""width=device-width,initial-scale=1"">
<title>Mouse Stability Report — {H(desc)}</title>
<style>
  body{{font-family:Segoe UI,Arial,sans-serif;margin:0;padding:0;background:#f5f5f5;color:#222}}
  main{{max-width:2000px;margin:0 auto;padding:24px}}
  h1{{font-size:2em;margin-bottom:4px}}
  h2{{font-size:1.4em;border-bottom:2px solid #bbb;padding-bottom:4px;margin-top:40px}}
  section{{background:#fff;border-radius:6px;padding:20px 24px;margin-bottom:24px;box-shadow:0 1px 4px #0002}}
  table{{border-collapse:collapse;width:100%;margin:8px 0}}
  th,td{{border:1px solid #ddd;padding:6px 12px;text-align:left}}
  th{{background:#f0f0f0}}
  .ok{{color:#2a7a2a;font-weight:bold}}
  .warn{{color:#b85000;font-weight:bold}}
  .risk{{color:#c00;font-weight:bold}}
  .meta{{color:#555;margin-bottom:0}}
  .v-excellent{{color:#1a6a1a;font-weight:bold}}
  .v-good{{color:#2a7a2a}}
  .v-acceptable{{color:#4a8a00}}
  .v-caution{{color:#b85000;font-weight:bold}}
  .v-risk{{color:#c00;font-weight:bold}}
  .phase-none{{color:#888}}
  .phase-minimal{{color:#2a7a2a}}
  .phase-slight{{color:#4a8a00}}
  .phase-moderate{{color:#b85000}}
  .phase-noticeable{{color:#c00;font-weight:bold}}
  .phase-critical{{color:#c00;font-weight:bold}}
  .legend{{background:#f8f8f8;border:1px solid #ddd;border-radius:4px;padding:12px 16px;margin-top:12px;font-size:.92em}}
  .legend dt{{font-weight:bold;margin-top:6px}}
  .legend dd{{margin:2px 0 0 16px;color:#444}}
  .tip{{position:relative;cursor:help;border-bottom:1px dotted #888;}}
  .tip::after{{content:attr(data-tip);position:absolute;left:0;top:calc(100% + 6px);background:#222;color:#f0f0f0;padding:8px 12px;border-radius:5px;font-size:.84em;width:340px;z-index:200;display:none;white-space:normal;line-height:1.55;pointer-events:none;box-shadow:0 3px 10px rgba(0,0,0,.35);}}
  .tip:hover::after{{display:block;}}
</style>
</head>
<body>
<main>
<h1>Mouse Stability Report</h1>
<p class=""meta""><strong>Device:</strong> {H(desc)} &nbsp;|&nbsp;
<strong>Detected:</strong> {r.DetectedHz} Hz &nbsp;|&nbsp;
<strong>Samples:</strong> {r.SampleCount} &nbsp;|&nbsp;
<strong>Generated:</strong> {DateTime.Now:yyyy-MM-dd HH:mm:ss}</p>
";
        }

        private static string SummaryTable(AnalysisResult r)
        {
            var sb = new StringBuilder();
            sb.Append("<table><tr><th>Metric</th><th>Value</th><th>Metric</th><th>Value</th></tr>");
            Row(sb, "Detected Hz",  $"{r.DetectedHz} Hz",
                    "Interval Nominal", $"{r.TNominal:F4} ms",
                tip1: "Polling rate inferred from the median interval. Nearest match from {125, 250, 500, 1000, 2000, 4000, 8000} Hz.",
                tip2: "Expected interval between polls at the detected rate: 1000 / Hz ms. Used as the baseline for all timing checks.");
            Row(sb, "Median Interval", $"{r.MedianInterval:F4} ms",
                    "Mean Interval",   $"{r.MeanInterval:F4} ms",
                tip1: "Middle value of all inter-poll intervals. Robust to outliers — a reliable proxy for the true polling period.",
                tip2: "Arithmetic mean of all inter-poll intervals. Should be close to the nominal interval for a well-behaved mouse.");
            Row(sb, "Std Dev",      $"{r.StddevInterval:F4} ms",
                    "CV (Jitter)",   $"{r.CV:F2}%  — {H(r.JitterRating)}",
                tip1: "Sample standard deviation of intervals. Measures how spread the timing is around the mean.",
                tip2: "Coefficient of Variation = StdDev / Mean × 100%. Relative timing variability independent of polling rate. < 5% Excellent · 5–15% Good · 15–30% Acceptable · > 30% Poor.");
            Row(sb, "Min Interval", $"{r.MinInterval:F4} ms",
                    "Max Interval",  $"{r.MaxInterval:F4} ms",
                tip1: "Shortest measured inter-poll interval. A value well below the nominal indicates back-to-back rapid polls.",
                tip2: "Longest measured inter-poll interval. A very large value indicates a severe drop or pause in reporting.");
            Row(sb, "Phase Jitter (measured): Avg", $"{r.MeanInterval:F4} ms",
                    "Phase Jitter (measured): P1",   $"{r.P1Interval:F4} ms",
                tip1: "Mean interval from captured data (same as Mean Interval above). Should be close to the nominal period. A noticeably higher value means the mouse polls slower than declared.",
                tip2: "1st-percentile interval — the fastest 1% of polls. A value well below the nominal indicates timing bunching: some polls arrive much earlier than expected, usually offset by equally late polls, causing micro-stutters.");
            Row(sb, "P99 Interval", $"{r.P99Interval:F4} ms",
                    "P99.9 Interval",$"{r.P999Interval:F4} ms",
                tip1: "99th-percentile interval. 99% of all polls arrive within this time. If this exceeds one display frame duration, occasional frame misses are likely.",
                tip2: "99.9th-percentile interval. Worst-case tail timing. If this exceeds one display frame, very rare but real missed frames will occur.");
            Row(sb, "Drop Count",   $"{r.DropCount} ({r.DropRate:F3}%)",
                    "Total Missed",  $"{r.TotalMissedPolls} polls",
                tip1: "Intervals exceeding 1.5× the nominal period. Each represents a delayed or skipped poll that can cause visible stuttering.",
                tip2: "Estimated total polls skipped across all drop events: sum of (round(interval / T_nominal) - 1) per drop.");
            Row(sb, "Null Packets", $"{r.NullTotal} ({r.NullRate:F2}%)",
                    "Null in Motion",$"{r.NullDuringMotion}",
                tip1: "Reports where both X and Y movement are zero. Can be normal firmware behavior or indicate sensor issues at high speed.",
                tip2: "Null packets flanked by non-null reports on both sides — the sensor stopped reporting during active movement.");
            Row(sb, "Dup Rate",     $"{r.DupRate:F2}%",
                    "Effective Hz",  $"{r.EffectiveHz:F0} Hz (×{r.FakeRatio:F2})",
                tip1: "Consecutive reports with identical non-zero coordinates, as a percentage. High values suggest duplicate data injection (fake polling).",
                tip2: "Polling rate derived from real (non-duplicate, non-null) intervals only. A large ratio vs. Detected Hz strongly suggests fake polling.");
            Row(sb, "Jump Events",  $"{r.JumpCount}",
                    "Jump Threshold",$"{r.JumpThreshold:F1} counts",
                tip1: "Reports where the displacement exceeded the jump threshold. May indicate tracking loss or sensor glitches.",
                tip2: "Displacement cutoff for jump detection: max(median × 10, mean + 5 × stddev) in raw sensor counts.");
            sb.Append("</table>");
            return sb.ToString();
        }

        private static string CompatTable(AnalysisResult r)
        {
            var sb = new StringBuilder();
            sb.Append("<table>");
            sb.Append("<tr>" +
                "<th><span class=\"tip\" data-tip=\"Monitor refresh rate being evaluated.\">Display</span></th>" +
                "<th><span class=\"tip\" data-tip=\"Duration of one display frame: 1000 / Hz ms. A polling interval exceeding this means no cursor update for that frame.\">Frame Time</span></th>" +
                "<th><span class=\"tip\" data-tip=\"Average number of mouse polls per display frame at this combination of polling rate and refresh rate. Values below 1 mean the cursor may not update every frame.\">Polls / Frame</span></th>" +
                "<th><span class=\"tip\" data-tip=\"Theoretical unevenness from a non-integer polls-per-frame ratio. When P/D is not a whole number some frames get one extra poll, causing slight cursor distance variation at constant speed. Formula: 1 / floor(P/D) × 100%. None = integer ratio · Minimal &lt;15% · Slight 15–30% · Moderate 30–50% · Noticeable ≥50% · Critical = &lt;1 poll/frame.\">Phase Jitter ①</span></th>" +
                "<th><span class=\"tip\" data-tip=\"Number of measured polling intervals from this session that exceeded one full display frame duration. Each such interval means the cursor did not move at all during that frame, causing visible input lag.\">Frame-Spanning Drops ② (measured)</span></th>" +
                "<th><span class=\"tip\" data-tip=\"Combined score starting at 4 (Excellent). Deductions: Phase Jitter (Slight −1, Moderate −2, Noticeable −3, Critical −4) + Frame-spanning drops (any −1, P99 exceeds frame −2, P99.9 exceeds frame −3) + Stability (CV&gt;15% or drop-rate&gt;1% → −1; CV&gt;30% or drop-rate&gt;5% → −2). Score 4=Excellent · 3=Good · 2=Acceptable · 1=Caution · 0=Risk.\">Verdict ③</span></th>" +
                "</tr>");

            foreach (var e in r.DisplayCompat)
            {
                // Polls/frame cell
                string pollsCell = e.PollsPerFrame >= 1
                    ? $"{e.PollsPerFrame:F2}"
                    : $"<span class=\"phase-critical\">{e.PollsPerFrame:F2} (&lt;1)</span>";

                // Phase jitter cell
                string phaseCls  = $"phase-{e.PhaseLabel.ToLower()}";
                string phaseCell = double.IsNaN(e.PhaseJitterPct)
                    ? $"<span class=\"{phaseCls}\">{H(e.PhaseLabel)}</span>"
                    : e.PhaseJitterPct == 0
                        ? $"<span class=\"{phaseCls}\">None (integer ratio)</span>"
                        : $"<span class=\"{phaseCls}\">{H(e.PhaseLabel)} ({e.PhaseJitterPct:F1}%)</span>";

                // Measured frame-spanning drops cell
                string dropCell = e.FrameSpanDrops == 0
                    ? "<span style=\"color:#2a7a2a\">0</span>"
                    : $"<span style=\"color:#c00\">{e.FrameSpanDrops} ({e.FrameDropPct:F3}%)</span>";

                // Verdict cell
                string verdictCls = e.Verdict == "Excellent"  ? "v-excellent"
                                  : e.Verdict == "Good"       ? "v-good"
                                  : e.Verdict == "Acceptable" ? "v-acceptable"
                                  : e.Verdict == "Caution"    ? "v-caution"
                                  :                             "v-risk";
                string verdictPrefix = e.Verdict is "Excellent" or "Good" or "Acceptable"
                    ? "✓" : e.Verdict == "Caution" ? "⚠" : "✗";

                sb.Append($"<tr>" +
                    $"<td>{e.DisplayHz} Hz</td>" +
                    $"<td>{e.FrameMs:F2} ms</td>" +
                    $"<td>{pollsCell}</td>" +
                    $"<td>{phaseCell}</td>" +
                    $"<td>{dropCell}</td>" +
                    $"<td class=\"{verdictCls}\">{verdictPrefix} {H(e.Verdict)}</td>" +
                    $"</tr>");
            }
            sb.Append("</table>");

            // Legend
            sb.Append(@"<dl class=""legend"">
<dt>① Phase Jitter</dt>
<dd>Inherent unevenness caused by a non-integer polls-per-frame ratio.
When P/D is not a whole number, some frames receive one more poll than others,
causing the cursor to advance a slightly different distance each frame even at constant velocity.
Metric = 1 / floor(P/D) × 100%: the maximum relative variation in per-frame displacement.
<br>Thresholds: <b>None</b> (integer ratio, 0%) · <b>Minimal</b> (&lt;15%) · <b>Slight</b> (15–30%) ·
<b>Moderate</b> (30–50%) · <b>Noticeable</b> (≥50%) · <b>Critical</b> (&lt;1 poll/frame).</dd>

<dt>② Frame-Spanning Drops (measured)</dt>
<dd>Number of measured polling intervals from this session that exceeded one full display frame.
A single such interval means the cursor did not move at all during that frame, causing visible input lag.
This is computed directly from the captured timing data.</dd>

<dt>③ Verdict</dt>
<dd>Combined score from three factors (each can deduct points from a starting score of 4):
<b>Phase Jitter</b> (Slight −1, Moderate −2, Noticeable −3, Critical −4) +
<b>Measured drops</b> (any frame-spanning drop −1, P99 exceeds frame −2, P99.9 exceeds frame −3) +
<b>Measured stability</b> (CV&gt;15% or drop-rate&gt;1% → −1; CV&gt;30% or drop-rate&gt;5% → −2).
Score 4=Excellent · 3=Good · 2=Acceptable · 1=Caution · 0=Risk.</dd>
</dl>");

            return sb.ToString();
        }

        private static void Row(StringBuilder sb,
            string k1, string v1, string k2, string v2,
            string tip1 = null, string tip2 = null)
        {
            string label1 = tip1 != null
                ? $"<span class=\"tip\" data-tip=\"{A(tip1)}\">{H(k1)}</span>"
                : H(k1);
            string label2 = tip2 != null
                ? $"<span class=\"tip\" data-tip=\"{A(tip2)}\">{H(k2)}</span>"
                : H(k2);
            sb.Append($"<tr><td>{label1}</td><td>{H(v1)}</td><td>{label2}</td><td>{H(v2)}</td></tr>");
        }

        private static string ToBase64(byte[] data)
            => Convert.ToBase64String(data);

        // HTML-encode a string (prevent XSS from device description)
        private static string H(string s)
            => s?.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;") ?? "";

        // Encode for use inside an HTML attribute value (double-quoted)
        private static string A(string s)
            => s?.Replace("&", "&amp;").Replace("\"", "&quot;") ?? "";
    }
}
