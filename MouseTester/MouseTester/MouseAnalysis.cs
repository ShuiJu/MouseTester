using System;
using System.Collections.Generic;
using System.Linq;

namespace MouseTester
{
    // ── Per-display-rate compatibility entry ───────────────────────────────────
    public class DisplayCompatEntry
    {
        public int    DisplayHz;
        public double FrameMs;

        // Ratio info
        public double PollsPerFrame;   // DetectedHz / DisplayHz (float)

        // Phase jitter: inherent unevenness from non-integer P/D ratio
        // = 1/floor(P/D) * 100 %, or NaN when P < D (less than 1 poll/frame)
        public double PhaseJitterPct;
        public string PhaseLabel;      // None / Minimal / Slight / Moderate / Noticeable / Critical

        // Measured: intervals from actual data that exceeded frame_ms
        public int    FrameSpanDrops;  // count
        public double FrameDropPct;    // % of total intervals

        // Combined verdict, driven by phase jitter + measured drops + overall stability
        public string Verdict;         // Excellent / Good / Acceptable / Caution / Risk
    }

    // ── Main result container ──────────────────────────────────────────────────
    public class AnalysisResult
    {
        // ── Range ──────────────────────────────────────────────────────────
        public int StartIdx, EndIdx, SampleCount;

        // ── Detected polling rate ──────────────────────────────────────────
        public int    DetectedHz;
        public double TNominal;          // expected interval in ms

        // ── Interval statistics ────────────────────────────────────────────
        public double MedianInterval, MeanInterval, StddevInterval, CV;
        public double MinInterval, MaxInterval, P1Interval, P99Interval, P999Interval;

        // ── Frame drops ────────────────────────────────────────────────────
        public int    DropCount;         // intervals exceeding 1.5 × T_nominal
        public int    TotalMissedPolls;  // sum of skipped polls
        public int    MaxConsecutiveMiss;
        public int    WorstDropIdx;      // event index at end of longest gap
        public double DropRate;          // %

        // ── Null packets ───────────────────────────────────────────────────
        public int    NullTotal;
        public int    NullDuringMotion;  // null sandwiched between non-null events
        public int    WorstNullCenterIdx;
        public double NullRate;          // %

        // ── Fake polling rate ──────────────────────────────────────────────
        public double DupRate;           // consecutive identical non-null pairs, %
        public double AltNullRate;       // alternating non-null/null pairs, %
        public double EffectiveHz;       // 1000 / median(real intervals)
        public double FakeRatio;         // DetectedHz / EffectiveHz (≥1.8 = suspicious)

        // ── Jumps ──────────────────────────────────────────────────────────
        public int    JumpCount;
        public int    WorstJumpIdx;
        public double JumpThreshold;

        // ── Jitter ─────────────────────────────────────────────────────────
        public string JitterRating;      // Excellent / Good / Acceptable / Poor

        // ── Display compatibility ──────────────────────────────────────────
        public List<DisplayCompatEntry> DisplayCompat;
    }

    // ── Analysis logic ─────────────────────────────────────────────────────────
    public static class MouseAnalysis
    {
        private static readonly int[] CandidateRates = { 125, 250, 500, 1000, 2000, 4000, 8000 };

        // All display refresh rates to evaluate
        private static readonly int[] DisplayRates =
            { 60, 120, 144, 165, 240, 360, 480, 540, 600, 750, 1000 };

        public static AnalysisResult Analyze(IList<MouseEvent> events, int startIdx, int endIdx)
        {
            var r = new AnalysisResult
            {
                StartIdx    = startIdx,
                EndIdx      = endIdx,
                SampleCount = endIdx - startIdx + 1
            };

            if (r.SampleCount < 3)
            {
                r.DetectedHz         = 1000;
                r.TNominal           = 1.0;
                r.JitterRating       = "N/A";
                r.WorstDropIdx       = startIdx;
                r.WorstNullCenterIdx = startIdx;
                r.WorstJumpIdx       = startIdx;
                r.DisplayCompat      = new List<DisplayCompatEntry>();
                return r;
            }

            int n = r.SampleCount;

            // ── Build intervals array ──────────────────────────────────────
            double[] intervals = new double[n - 1];
            for (int i = 1; i < n; i++)
                intervals[i - 1] = events[startIdx + i].ts - events[startIdx + i - 1].ts;

            // ── Detect polling rate ────────────────────────────────────────
            double medianInterval = Median(intervals);
            r.DetectedHz = CandidateRates
                .OrderBy(c => Math.Abs(1000.0 / c - medianInterval))
                .First();
            r.TNominal = 1000.0 / r.DetectedHz;

            // ── Interval statistics ────────────────────────────────────────
            r.MedianInterval  = medianInterval;
            r.MeanInterval    = intervals.Average();
            r.StddevInterval  = SampleStdDev(intervals, r.MeanInterval);
            r.CV              = r.MeanInterval > 0 ? r.StddevInterval / r.MeanInterval * 100.0 : 0;
            r.MinInterval     = intervals.Min();
            r.MaxInterval     = intervals.Max();
            r.P1Interval      = Percentile(intervals, 1.0);
            r.P99Interval     = Percentile(intervals, 99.0);
            r.P999Interval    = Percentile(intervals, 99.9);

            r.JitterRating = r.CV < 5  ? "Excellent"
                           : r.CV < 15 ? "Good"
                           : r.CV < 30 ? "Acceptable"
                           : "Poor";

            // ── Frame drops ────────────────────────────────────────────────
            double dropThreshold = r.TNominal * 1.5;
            int consecutive = 0, maxConsecutive = 0, worstIntervalIdx = 0;

            for (int i = 0; i < intervals.Length; i++)
            {
                if (intervals[i] > intervals[worstIntervalIdx]) worstIntervalIdx = i;

                if (intervals[i] > dropThreshold)
                {
                    r.DropCount++;
                    r.TotalMissedPolls += Math.Max(0, (int)Math.Round(intervals[i] / r.TNominal) - 1);
                    consecutive++;
                    if (consecutive > maxConsecutive) maxConsecutive = consecutive;
                }
                else
                {
                    consecutive = 0;
                }
            }
            r.MaxConsecutiveMiss = maxConsecutive;
            r.DropRate           = (double)r.DropCount / intervals.Length * 100.0;
            r.WorstDropIdx       = startIdx + worstIntervalIdx + 1;

            // ── Null packets ───────────────────────────────────────────────
            int worstNullDensity = 0;
            r.WorstNullCenterIdx = startIdx;

            for (int i = 0; i < n; i++)
            {
                var ev = events[startIdx + i];
                if (ev.lastx != 0 || ev.lasty != 0) continue;
                r.NullTotal++;

                if (i > 0 && i < n - 1)
                {
                    var prev = events[startIdx + i - 1];
                    var next = events[startIdx + i + 1];
                    if ((prev.lastx != 0 || prev.lasty != 0) &&
                        (next.lastx != 0 || next.lasty != 0))
                        r.NullDuringMotion++;
                }

                double center = ev.ts;
                int density = 0;
                for (int j = i; j >= 0 && events[startIdx + j].ts >= center - 5.0; j--)
                    if (events[startIdx + j].lastx == 0 && events[startIdx + j].lasty == 0) density++;
                for (int j = i + 1; j < n && events[startIdx + j].ts <= center + 5.0; j++)
                    if (events[startIdx + j].lastx == 0 && events[startIdx + j].lasty == 0) density++;

                if (density > worstNullDensity)
                {
                    worstNullDensity     = density;
                    r.WorstNullCenterIdx = startIdx + i;
                }
            }
            r.NullRate = (double)r.NullTotal / n * 100.0;

            // ── Fake polling detection ─────────────────────────────────────
            int dupCount = 0, altNullPairs = 0;
            var effIntervals = new List<double>(n);

            for (int i = 1; i < n; i++)
            {
                var cur  = events[startIdx + i];
                var prev = events[startIdx + i - 1];
                bool curNull  = cur.lastx  == 0 && cur.lasty  == 0;
                bool prevNull = prev.lastx == 0 && prev.lasty == 0;

                if (!curNull && cur.lastx == prev.lastx && cur.lasty == prev.lasty)
                    dupCount++;
                if (!prevNull && curNull)
                    altNullPairs++;
                if (!curNull && !(cur.lastx == prev.lastx && cur.lasty == prev.lasty))
                    effIntervals.Add(intervals[i - 1]);
            }
            r.DupRate     = (double)dupCount / (n - 1) * 100.0;
            r.AltNullRate = (double)altNullPairs / (n - 1) * 100.0;
            r.EffectiveHz = effIntervals.Count > 5
                ? 1000.0 / Median(effIntervals.ToArray())
                : r.DetectedHz;
            r.FakeRatio   = r.EffectiveHz > 0 ? r.DetectedHz / r.EffectiveHz : 1.0;

            // ── Jump detection ─────────────────────────────────────────────
            double[] displacements = new double[n];
            for (int i = 0; i < n; i++)
            {
                var ev = events[startIdx + i];
                displacements[i] = Math.Sqrt((double)ev.lastx * ev.lastx + (double)ev.lasty * ev.lasty);
            }
            var nonzero = displacements.Where(d => d > 0).ToArray();
            if (nonzero.Length > 0)
            {
                double medDisp  = Median(nonzero);
                double meanDisp = nonzero.Average();
                double stdDisp  = SampleStdDev(nonzero, meanDisp);
                r.JumpThreshold = Math.Max(medDisp * 10.0, meanDisp + 5.0 * stdDisp);
            }
            else
            {
                r.JumpThreshold = double.MaxValue;
            }

            int worstJumpLocal = 0;
            for (int i = 0; i < n; i++)
            {
                if (displacements[i] > displacements[worstJumpLocal]) worstJumpLocal = i;
                if (displacements[i] > r.JumpThreshold) r.JumpCount++;
            }
            r.WorstJumpIdx = startIdx + worstJumpLocal;

            // ── Display compatibility ──────────────────────────────────────
            r.DisplayCompat = new List<DisplayCompatEntry>(DisplayRates.Length);
            foreach (int dHz in DisplayRates)
            {
                double framems       = 1000.0 / dHz;
                double pollsPerFrame = (double)r.DetectedHz / dHz;

                // Phase jitter: inherent unevenness of poll-to-frame alignment
                double phaseJitterPct;
                string phaseLabel;
                if (pollsPerFrame < 1.0)
                {
                    phaseJitterPct = double.NaN; // < 1 poll/frame
                    phaseLabel     = "Critical";
                }
                else
                {
                    double frac = pollsPerFrame - Math.Floor(pollsPerFrame);
                    if (frac < 0.01)              // effectively an integer ratio
                    {
                        phaseJitterPct = 0;
                        phaseLabel     = "None";
                    }
                    else
                    {
                        // Relative variation = 1 missed poll ÷ polls per frame
                        phaseJitterPct = 1.0 / Math.Floor(pollsPerFrame) * 100.0;
                        phaseLabel     = phaseJitterPct >= 50 ? "Noticeable"
                                       : phaseJitterPct >= 30 ? "Moderate"
                                       : phaseJitterPct >= 15 ? "Slight"
                                       :                        "Minimal";
                    }
                }

                // Measured: how many actual intervals exceeded one frame duration
                int frameSpanDrops = 0;
                foreach (double iv in intervals)
                    if (iv > framems) frameSpanDrops++;
                double frameDropPct = (double)frameSpanDrops / intervals.Length * 100.0;

                // Combined verdict score (starts at 4 = Excellent)
                int score = 4;

                // Phase jitter deduction
                if (double.IsNaN(phaseJitterPct))   score -= 4; // Critical
                else if (phaseJitterPct >= 50)       score -= 3; // Noticeable
                else if (phaseJitterPct >= 30)       score -= 2; // Moderate
                else if (phaseJitterPct >= 15)       score -= 1; // Slight

                // Measured frame-spanning drops deduction
                if (r.P999Interval >= framems)       score -= 3;
                else if (r.P99Interval >= framems)   score -= 2;
                else if (frameDropPct > 0)           score -= 1;

                // Overall measured stability deduction
                if      (r.CV > 30 || r.DropRate > 5)  score -= 2;
                else if (r.CV > 15 || r.DropRate > 1)  score -= 1;

                score = Math.Max(0, score);
                string verdict = score >= 4 ? "Excellent"
                               : score >= 3 ? "Good"
                               : score >= 2 ? "Acceptable"
                               : score >= 1 ? "Caution"
                               :              "Risk";

                r.DisplayCompat.Add(new DisplayCompatEntry
                {
                    DisplayHz      = dHz,
                    FrameMs        = framems,
                    PollsPerFrame  = pollsPerFrame,
                    PhaseJitterPct = phaseJitterPct,
                    PhaseLabel     = phaseLabel,
                    FrameSpanDrops = frameSpanDrops,
                    FrameDropPct   = frameDropPct,
                    Verdict        = verdict
                });
            }

            return r;
        }

        // ── Window helper ──────────────────────────────────────────────────────
        public static (int s, int e) GetWindowIndices(
            IList<MouseEvent> events, int centerIdx, double halfWindowMs,
            int analysisStart, int analysisEnd)
        {
            double lo = events[centerIdx].ts - halfWindowMs;
            double hi = events[centerIdx].ts + halfWindowMs;
            int s = centerIdx;
            while (s > analysisStart && events[s - 1].ts >= lo) s--;
            int e = centerIdx;
            while (e < analysisEnd  && events[e + 1].ts <= hi) e++;
            return (s, e);
        }

        // ── Statistics helpers ─────────────────────────────────────────────────
        private static double Median(double[] values)
        {
            var sorted = (double[])values.Clone();
            Array.Sort(sorted);
            int mid = sorted.Length / 2;
            return sorted.Length % 2 == 0
                ? (sorted[mid - 1] + sorted[mid]) / 2.0
                : sorted[mid];
        }

        private static double SampleStdDev(double[] values, double mean)
        {
            if (values.Length < 2) return 0;
            double sum = values.Sum(v => (v - mean) * (v - mean));
            return Math.Sqrt(sum / (values.Length - 1));
        }

        private static double Percentile(double[] values, double pct)
        {
            var sorted = (double[])values.Clone();
            Array.Sort(sorted);
            double idx = pct / 100.0 * (sorted.Length - 1);
            int lo = (int)Math.Floor(idx);
            int hi = Math.Min((int)Math.Ceiling(idx), sorted.Length - 1);
            if (lo == hi) return sorted[lo];
            return sorted[lo] + (sorted[hi] - sorted[lo]) * (idx - lo);
        }
    }
}
