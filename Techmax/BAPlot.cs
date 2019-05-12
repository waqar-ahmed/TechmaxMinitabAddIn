using System;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

namespace Techmax
{
    internal partial class BAPlot : Form
    {
        internal BAPlot()
        {
            
            InitializeComponent();
        }

        internal BAPlot(ref Mtb.Application pApp)
        {
            
            AppDomain.CurrentDomain.AssemblyResolve += new ResolveEventHandler(CurrentDomain_AssemblyResolve);

            InitializeComponent();
            AddIn.gMtbApp = pApp;
        }

        static Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
        {
            return EmbeddedAssembly.Get(args.Name);
        }

        static byte[] StreamToBytes(Stream input)
        {
            var capacity = input.CanSeek ? (int)input.Length : 0;
            using (var output = new MemoryStream(capacity))
            {
                int readLength;
                var buffer = new byte[4096];

                do
                {
                    readLength = input.Read(buffer, 0, buffer.Length);
                    output.Write(buffer, 0, readLength);
                }
                while (readLength != 0);

                return output.ToArray();
            }
        }
    }
}
