using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SpreadsheetGUI
{
    public partial class TwitterEnableDialog : Form
    {
        public string userAlias;
        public TwitterEnableDialog()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            userAlias = aliasTextBox.Text;
        }
    }
}
