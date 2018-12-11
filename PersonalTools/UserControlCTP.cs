using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace peteli.PersonalTools
{
    [ComVisible(true)] //this is important excel throws expception otherwise!
    public partial class UserControlCTP : UserControl
    {
        public UserControlCTP()
        {
            InitializeComponent();
            this.pictureBox1.DoubleClick += PickFile;
        }

        private void PickFile(object sender, EventArgs e)
        {
            this.openFileDialogLogo.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Recent);
            if (this.openFileDialogLogo.ShowDialog() == DialogResult.OK)
            {
                this.pictureBox1.Image = new Bitmap(openFileDialogLogo.FileName);
            }
        }
    }
}
