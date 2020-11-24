using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GlasshouseExcel.View
{
    public partial class WaitingForm : Form
    {
        string _format;
        public WaitingForm(string caption, string format)
        {
            _format = format;
            InitializeComponent();
            Text = caption;
            label1.Text = (null == format) ? caption : string.Format(format, 0);
            pictureBox1.Image = Properties.Resources.waiting;
            Show();
            Application.DoEvents();
        }
    }
}
