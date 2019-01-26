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
using System.IO;

namespace peteli.PersonalTools
{
    [ComVisible(true)] //this is important excel throws expception otherwise!
    public partial class UserControlCTP : UserControl
    {
        // constructor
        public UserControlCTP()
        {
            InitializeComponent();
            this.openFileDialogLogo.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures);
           //this.openFileDialogLogo.FileName = "";
            this.pictureBox1.Click += PickFile;
            this.buttonFormat.Click += CorporateFormatSheet;
            this.buttonProperties.Click += ShowWorkbookProperties;
            this.Load += AssignLastLogo;
            this.Load += InitConfidentiality;
            this.Load += DataBindingWorkbookProperties;
        }

        private void DataBindingWorkbookProperties(object sender, EventArgs e)
        {
            
            // bind controls to class that holds workbook properties
            this.textBoxTitle.DataBindings.Add(new Binding("Text", this.docProps, "TitleString"));
            this.textBoxSubject.DataBindings.Add(new Binding("Text", docProps, "SubjectString"));
            this.textBoxAuthor.DataBindings.Add(new Binding("Text", this.docProps, "AuthorString"));
            this.textBoxManager.DataBindings.Add(new Binding("Text", this.docProps, "ManagerString"));
            this.textBoxCompany.DataBindings.Add(new Binding("Text", this.docProps, "CompanyString"));
            this.textBoxConfidentiality.DataBindings.Add(new Binding("Text", this.docProps, "ConfindentialityString"));
        }

        System.Collections.Specialized.StringCollection confLevel = new System.Collections.Specialized.StringCollection();
        WorkbookProperties docProps = new WorkbookProperties();

        private void AssignLastLogo(object sender, EventArgs e)
        {
            if (File.Exists(Properties.Settings.Default.LogoImagePath))
            {
                // last saved logo image path is valid
                // load picture 
                this.pictureBox1.Load(Properties.Settings.Default.LogoImagePath);
                this.openFileDialogLogo.FileName = Properties.Settings.Default.LogoImagePath;
            }
        }

        private void ShowWorkbookProperties(object sender, EventArgs e)
        {
            FormatHeaderFooter.ShowPropertyDialog();
        }

        private void CorporateFormatSheet(object sender, EventArgs e)
        {
             FormatHeaderFooter.Do(this.docProps);
        }

        private void PickFile(object sender, EventArgs e)
        {
            if (this.openFileDialogLogo.ShowDialog() == DialogResult.OK)
            {
                this.pictureBox1.Image = new Bitmap(openFileDialogLogo.FileName);
                Properties.Settings.Default.LogoImagePath = openFileDialogLogo.FileName;
                Properties.Settings.Default.Save();
            }
        }
        private void InitConfidentiality(object sender, EventArgs e)
        {
            String[] myArr = new String[] { "public", "for internal use only", "confidential"};
            confLevel.AddRange(myArr);
            //this.listBoxConfidentiality.DataSource = confLevel;
        }

    }
}
