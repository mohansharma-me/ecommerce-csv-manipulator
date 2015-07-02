using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace CSVManip
{
    public partial class frmWait : Form
    {
        public frmWait(String title, String msg)
        {
            InitializeComponent();
            this.lblTitle.Text = title;
            this.lblMsg.Text = msg;
            Text = title;
        }

        private void frmWait_Load(object sender, EventArgs e)
        {

        }
    }
}
