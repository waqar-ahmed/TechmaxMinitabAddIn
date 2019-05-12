using System.Windows.Forms;

namespace Techmax
{
    internal partial class FormRename : Form
    {
        internal FormRename()
        {
            InitializeComponent();
        }

        internal FormRename(ref Mtb.Application pApp)
        {
            InitializeComponent();
            AddIn.gMtbApp = pApp;
        }
    }
}
