using OxyPlot;
using OxyPlot.Series;
using Python.Runtime;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using ZedGraph;

namespace Techmax
{
    internal partial class BAPlot : Form
    {
        internal BAPlot(ref Mtb.Application pApp)
        {
            AppDomain.CurrentDomain.AssemblyResolve += new ResolveEventHandler(CurrentDomain_AssemblyResolve);
            InitializeComponent();
            AddIn.gMtbApp = pApp;
           

            //var myModel = new PlotModel { Title = "Example 1" };
            //myModel.Series.Add(new FunctionSeries(Math.Cos, 0, 10, 0.1, "cos(x)"));
            //this.plot1.Model = myModel;
        }

        static Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
        {
            return EmbeddedAssembly.Get(args.Name);
        }

        private void BAPlot_Load(object sender, EventArgs e)
        {
            z1.IsShowPointValues = true;
            z1.GraphPane.Title = "Test Case for C#";
            double[] x = new double[100];
            double[] y = new double[100];
            int i;
            for (i = 0; i < 100; i++)
            {
                x[i] = (double)i / 100.0 * Math.PI * 2.0;
                y[i] = Math.Sin(x[i]);
            }
            z1.GraphPane.AddCurve("Sine Wave", x, y, Color.Red, SymbolType.Square);
            z1.AxisChange();
            z1.Invalidate();
        }
    }
}
