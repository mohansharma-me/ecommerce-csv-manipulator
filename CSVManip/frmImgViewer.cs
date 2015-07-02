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
    public partial class frmImgViewer : Form
    {
        public frmImgViewer(Image img)
        {
            InitializeComponent();
            this.pictureBox1.Image = img;
        }

        private void frmImgViewer_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
