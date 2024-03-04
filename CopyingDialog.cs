using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EmulatorStarter
{
    public partial class CopyingDialog : Form
    {
        public CopyingDialog()
        {
            InitializeComponent();

            this.FormClosing += CopyingDialog_FormClosing;
        }

        private void CopyingDialog_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
            }
        }
    }
}
