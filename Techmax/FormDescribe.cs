﻿using System.Windows.Forms;

namespace Techmax
{
    internal partial class FormDescribe : Form
    {
        internal FormDescribe()
        {
            InitializeComponent();
        }

        internal FormDescribe(ref Mtb.Application pApp)
        {
            InitializeComponent();
            AddIn.gMtbApp = pApp;
        }

        private void FormDescribe_Load(object sender, System.EventArgs e)
        {

        }
    }
}
