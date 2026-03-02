using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

namespace MouseTester
{
    public class MouseLog
    {
        private readonly object _lock = new object();
        private string desc = "MouseTester";
        private double cpi = 400.0;
        private List<MouseEvent> events = new List<MouseEvent>();

        public double Cpi
        {
            get { return this.cpi; }
            set { cpi = value; }
        }

        public string Desc
        {
            get { return this.desc; }
            set { this.desc = value; }
        }

        public List<MouseEvent> Events
        {
            get { return this.events; }
        }

        public void Add(MouseEvent e)
        {
            lock (_lock) { this.events.Add(e); }
        }

        public void Clear()
        {
            lock (_lock) { this.events.Clear(); }
        }

        public void Load(string fname)
        {
            var loaded = new List<MouseEvent>();
            string newDesc;
            double newCpi;

            using (StreamReader sr = File.OpenText(fname))
            {
                newDesc = sr.ReadLine();
                if (!double.TryParse(sr.ReadLine(), NumberStyles.Any, CultureInfo.InvariantCulture, out newCpi))
                    throw new FormatException("Invalid CPI value in file.");

                string headerline = sr.ReadLine();
                while (sr.Peek() > -1)
                {
                    string line = sr.ReadLine();
                    string[] values = line.Split(',');
                    if (values.Length == 4)
                    {
                        loaded.Add(new MouseEvent(
                            ushort.Parse(values[3]),
                            int.Parse(values[0]),
                            int.Parse(values[1]),
                            double.Parse(values[2], CultureInfo.InvariantCulture)));
                    }
                    else if (values.Length == 3)
                    {
                        loaded.Add(new MouseEvent(
                            0,
                            int.Parse(values[0]),
                            int.Parse(values[1]),
                            double.Parse(values[2], CultureInfo.InvariantCulture)));
                    }
                }
            }

            lock (_lock)
            {
                this.desc = newDesc;
                this.cpi = newCpi;
                this.events = loaded;
            }
        }

        public void Save(string fname)
        {
            List<MouseEvent> snapshot;
            string desc;
            double cpi;

            lock (_lock)
            {
                snapshot = new List<MouseEvent>(this.events);
                desc = this.desc;
                cpi = this.cpi;
            }

            using (StreamWriter sw = File.CreateText(fname))
            {
                sw.WriteLine(desc);
                sw.WriteLine(cpi.ToString(CultureInfo.InvariantCulture));
                sw.WriteLine("xCount,yCount,Time (ms),buttonflags");
                foreach (MouseEvent e in snapshot)
                {
                    sw.WriteLine(e.lastx.ToString() + "," + e.lasty.ToString() + "," +
                                 e.ts.ToString(CultureInfo.InvariantCulture) + "," + e.buttonflags.ToString());
                }
            }
        }

        public int deltaX()
        {
            return this.events.Sum(e => e.lastx);
        }

        public int deltaY()
        {
            return this.events.Sum(e => e.lasty);
        }

        public double path()
        {
            return this.events.Sum(e => Math.Sqrt((e.lastx * e.lastx) + (e.lasty * e.lasty)));
        }
    }
}
