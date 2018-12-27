using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace peteli.PersonalTools
{
    /// <summary>
    /// Interaction logic for UserControlCTPxaml.xaml
    /// </summary>
    public partial class UserControlCTPxaml : UserControl
    {
        #region contructor
        public UserControlCTPxaml()
        {
            InitializeComponent();
            this.Loaded += BindWorkbookProperties;
            this.Loaded += AssignLastLogoOrDefault;
            this.imgLogo.MouseLeftButtonDown += PickImageFile;
            this.cmdbtnFormat.Click += CorporateFormatSheet;
            this.btnLoadDefaults.Click += BtnLoadDefaults_Click;
            this.btnSaveDefaults.Click += BtnSaveDefaults_Click;
        }
        #endregion

        private void BtnSaveDefaults_Click(object sender, RoutedEventArgs e)
        {
            WorkbookProperties.UserDefaults.AuthorString = docProps.AuthorString;
            WorkbookProperties.UserDefaults.ManagerString = docProps.ManagerString;
            WorkbookProperties.UserDefaults.CompanyString = docProps.CompanyString;
            WorkbookProperties.UserDefaults.ConfindentialityString = docProps.ConfindentialityString;
            WorkbookProperties.UserDefaults.Save();
        }

        private void BtnLoadDefaults_Click(object sender, RoutedEventArgs e)
        {
            docProps.AuthorString = WorkbookProperties.UserDefaults.AuthorString;
            docProps.ManagerString = WorkbookProperties.UserDefaults.ManagerString;
            docProps.CompanyString = WorkbookProperties.UserDefaults.CompanyString;
            docProps.ConfindentialityString = WorkbookProperties.UserDefaults.ConfindentialityString;
        }

        private void CorporateFormatSheet(object sender, RoutedEventArgs e)
        {
            FormatHeaderFooter.Do(this.docProps);
        }

        private void PickImageFile(object sender, MouseButtonEventArgs e)
        {
            if (this.imagePicker.ShowDialog() == true)
            {
                this.imgLogo.Source = new BitmapImage(new Uri(imagePicker.FileName));
                docProps.LogoImageFileName = imagePicker.FileName;
                //Properties.Settings.Default.Save();
            }
        }

        private void AssignLastLogoOrDefault(object sender, RoutedEventArgs e)
        {
            if (File.Exists(docProps.LogoImageFileName))
            {
                // last saved logo image path is valid
                // load picture 
                this.imgLogo.Source = new BitmapImage(new Uri(docProps.LogoImageFileName));
                this.imagePicker.FileName = docProps.LogoImageFileName;
            }
            else
            {
                ImageSourceConverter c = new ImageSourceConverter();
                this.imgLogo.Source = (ImageSource)c.ConvertFrom(Properties.Resources.logo_default);
                this.imagePicker.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures);
            }
            this.imgLogo.Stretch = Stretch.Uniform;

        }

        private void BindWorkbookProperties(object sender, EventArgs e)
        {
            // bind controls to class that holds workbook properties
            this.DataContext = this.docProps;
            this.bxTitle.SetBinding(TextBox.TextProperty, new Binding("TitleString")
            {
                Mode = BindingMode.TwoWay
            } );
            this.bxSubject.SetBinding(TextBox.TextProperty, "SubjectString");
            this.bxAuthor.SetBinding(TextBox.TextProperty, "AuthorString");
            this.bxManager.SetBinding(TextBox.TextProperty, "ManagerString");
            this.bxCompany.SetBinding(TextBox.TextProperty, "CompanyString");
            this.bxConfidentiality.SetBinding(TextBox.TextProperty, new Binding("ConfindentialityString") { Mode= BindingMode.TwoWay});
        }

        #region properties
        WorkbookProperties docProps = new WorkbookProperties();
        Microsoft.Win32.OpenFileDialog imagePicker = new Microsoft.Win32.OpenFileDialog()
        {
            Filter = "Images|*.jpg;*.jpeg;*.bmp;*.png;*.gif;*.tiff"
        };
        #endregion
    }
}
