using System;
using System.Xml;
using System.Xml.Linq;
using System.IO;
using System.Collections.Generic;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using System.Text;
using QRCoder;
using System.Drawing;
using System.ComponentModel;

namespace peteli.PersonalTools
{
    /// <summary>
    /// class holding model to modify header/footer of excel workbook
    /// </summary>
    static class FormatHeaderFooter
    {
        internal static void ShowPropertyDialog()
        {
            Application xlApp = (Application)ExcelDnaUtil.Application;
            Dialog xlSummaryDialog = xlApp.Dialogs.Item[XlBuiltInDialog.xlDialogSummaryInfo];
            var result = xlSummaryDialog.Show();
        }

        internal static void Do(WorkbookProperties doc_props)
        {
            // create instance
            WorkbookProperties docProps = doc_props;

            Application XlApp = (Application)ExcelDnaUtil.Application;
            Worksheet Xlws = XlApp.ActiveWindow.ActiveSheet;

            Xlws.PageSetup.ScaleWithDocHeaderFooter = false;
            Xlws.PageSetup.AlignMarginsHeaderFooter = true;

            if (Xlws.Equals(null)) return;


            #region RightHeaderGrafic
            Graphic imgRightHeader = Xlws.PageSetup.RightHeaderPicture;
            imgRightHeader.Filename = docProps.LogoImageFileName;
            imgRightHeader.LockAspectRatio =  Microsoft.Office.Core.MsoTriState.msoTrue;
            imgRightHeader.Height = 25;
            //imgRightHeader.Width = 463.5;
            //imgRightHeader.Brightness = 0.36;
            //imgRightHeader.ColorType = msoPictureGrayscale;
            //imgRightHeader.Contrast = 0.39;
            //imgRightHeader.CropBottom = -14.4;
            //imgRightHeader.CropLeft = -28.8;
            //imgRightHeader.CropRight = -14.4;
            //imgRightHeader.CropTop = 21.6;
            #endregion
            #region LeftFooterGrafic
            // didn't find a way other then saving QRcode image first and secondly take filname and assign it to graphic object
            string imageFileName = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
            docProps.GetQRCodeImage(docProps.GetXMLString).Save(imageFileName);
            Graphic imgLeftFooter = Xlws.PageSetup.LeftFooterPicture;
            imgLeftFooter.Filename = imageFileName;
            imgLeftFooter.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;
            imgLeftFooter.Height = 60;

            #endregion
            #region apply changes
            Xlws.PageSetup.CenterHeader = docProps.CenterHeader.ToString();
            Xlws.PageSetup.LeftHeader = docProps.LeftHeader.ToString();
            Xlws.PageSetup.RightHeader = docProps.RightHeader.ToString();

            // if char in total header and footer are more than 255 than do grafic instead of text
            if (docProps.IsToManyCharactersHeaderFooter)
            {
                Xlws.PageSetup.LeftFooter = docProps.LeftFooterGrafic.ToString();
            }
            else
            {
                Xlws.PageSetup.LeftFooter = docProps.LeftFooter.ToString();
            }
            Xlws.PageSetup.CenterFooter = docProps.CenterFooter.ToString();
            Xlws.PageSetup.RightFooter = docProps.RightFooter.ToString();
            #endregion
        }
    }
    public class WorkbookProperties : INotifyPropertyChanged
    {
        #region event declaration
        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string propertyName){PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));}
        #endregion

        #region properties
        static Application _XlApp = (Application)ExcelDnaUtil.Application;
        DocumentProperty Title { get { return _XlApp.ActiveWorkbook.BuiltinDocumentProperties(nameof(Title)); } }
        DocumentProperty Subject { get { return _XlApp.ActiveWorkbook.BuiltinDocumentProperties(nameof(Subject)); } }
        DocumentProperty Author { get { return _XlApp.ActiveWorkbook.BuiltinDocumentProperties(nameof(Author)); } }
        DocumentProperty Manager { get { return _XlApp.ActiveWorkbook.BuiltinDocumentProperties(nameof(Manager)); } }
        DocumentProperty Company { get { return _XlApp.ActiveWorkbook.BuiltinDocumentProperties(nameof(Company)); } }
        DocumentProperty Confidentiality
        { get
            {
                try
                {
                    return _XlApp.ActiveWorkbook.CustomDocumentProperties(nameof(Confidentiality));
                }
                catch
                {
                    DocumentProperties docProps = _XlApp.ActiveWorkbook.CustomDocumentProperties;
                    DocumentProperty newConfidentiality =
                        docProps.Add(nameof(Confidentiality), false, MsoDocProperties.msoPropertyTypeString, UserDefaults.ConfindentialityString);
                    //newConfidentiality.Value = UserDefaults.ConfindentialityString;
                    return newConfidentiality;
                }
            }
        }
        public string TitleString { get { return (string)Title.Value; } set { Title.Value = value; } }
        public string SubjectString { get { return (string)Subject.Value; } set { Subject.Value = value; } }
        public string AuthorString { get { return (string)Author.Value; }
            set
            {
                Author.Value = value;
                // Call OnPropertyChanged whenever the property is updated for Binding
                OnPropertyChanged(nameof(AuthorString));
            }
        }
        public string ManagerString { get { return (string)Manager.Value; }
            set
            {
                Manager.Value = value;
                // Call OnPropertyChanged whenever the property is updated for Binding
                OnPropertyChanged(nameof(ManagerString));
            }
        }
        public string CompanyString { get { return (string)Company.Value; }
            set
            {
                Company.Value = value;
                // Call OnPropertyChanged whenever the property is updated for Binding
                OnPropertyChanged(nameof(CompanyString));
            }
        }
        public string ConfindentialityString { get { return (string)Confidentiality.Value; }
            set
            {
                Confidentiality.Value = value;
                // Call OnPropertyChanged whenever the property is updated for Binding
                OnPropertyChanged(nameof(ConfindentialityString));
            } }
        public string LogoImageFileName { get { return UserDefaults.LogoImageFileName; }
            set
            {
                UserDefaults.LogoImageFileName = value;
                // Call OnPropertyChanged whenever the property is updated for Binding
                OnPropertyChanged(nameof(LogoImageFileName));
            } }
        public string FontName =>  new StringBuilder()
                    .Append("&\"")
                    .Append(Properties.Settings.Default.FontName)
                    .Append("\"")
                    .Append("&6")
                    .ToString();
        public string FontNameConfidentiality => new StringBuilder()
            .Append("&\"")
            .Append(Properties.Settings.Default.FontNameConfidentiality)
            .Append("\"")
            .Append("&6")
            .ToString();
        public char SpecialAttention => Properties.Settings.Default.charAttention;

        public StringBuilder LeftHeader => new StringBuilder()
            .Append(FontName)
            .Append("&10&B")
            .Append(TitleString + "\n")
            //.Append()
            .Append("&B")
            .Append(SubjectString)
            .Append("&10");
        public StringBuilder CenterHeader => new StringBuilder()
            .Append(FontName)
            .Append(_XlApp.ActiveWindow.ActiveSheet.Name);
        public StringBuilder RightHeader => RightHeaderGrafic;
        public StringBuilder RightHeaderGrafic => new StringBuilder("&G");
        public StringBuilder LeftFooterGrafic => new StringBuilder("&G");
        public StringBuilder LeftFooter => new StringBuilder()
            .Append(FontName)
            .Append("Author: " + AuthorString)
            .Append("\n")
            .Append("Manager: " + ManagerString)
            .Append("\n")
            .Append(CompanyString);
        public StringBuilder LeftFooterQRCodeString => new StringBuilder()
            .Append("Author: " + AuthorString)
            .Append("Manager: " + ManagerString)
            .Append("Company: " + CompanyString)
            .Append(_XlApp.ActiveWorkbook.FullName);
        public StringBuilder CenterFooter => new StringBuilder()
            .Append(FontName)
            .Append("Page &P of &N")
            .Append("\n")
            .Append(FontNameConfidentiality)
            .Append(SpecialAttention).Append(" ")
            .Append("&B")
            .Append(ConfindentialityString)
            .Append("&B")
            .Append(" ").Append(SpecialAttention);
        public StringBuilder RightFooter => new StringBuilder()
            .Append(FontName)
            .Append("&F")
            .Append("\n")
            .Append("Printed: &D &T");
        public Graphic ImageLogo { get; set; }
        public const byte MaxCharactersHeaderFooter = 200;
        public Int32 ActualCharactersHeaderFooter => LeftFooter.Length + CenterFooter.Length + RightFooter.Length;
        public bool IsToManyCharactersHeaderFooter => (ActualCharactersHeaderFooter >= MaxCharactersHeaderFooter) ? true : false;
        #endregion
        #region QRcode
        public Bitmap GetQRCodeImage(string QRCodeText)
        {
            QRCodeGenerator qrGenerator = new QRCodeGenerator();
            QRCodeData qrCodeData = qrGenerator.CreateQrCode(QRCodeText, QRCodeGenerator.ECCLevel.Q);
            QRCode qrCode = new QRCode(qrCodeData);
            Bitmap qrCodeImage = qrCode.GetGraphic(20);

            return qrCodeImage;
        }
        #endregion
        #region XMLdocument
        internal XDocument GetXMLDoc()
        {
            XDocument doc = new XDocument(
                new XDeclaration("1.0", "utf-8", string.Empty),
                new XElement("workbook",
                    new XElement("filename", _XlApp.ActiveWorkbook.FullName),
                    new XElement("author", AuthorString),
                    new XElement("manager",ManagerString),
                    new XElement("company",CompanyString)
                ));
            Console.Write(doc);
            return doc;
        }
        internal XDocument GetHTML()
        {
            XDocument doc = new XDocument(
             new XDocumentType("html", null, null, null),
             new XElement("html",
                new XElement("head"),
                    new XElement("body",
                    new XElement("p", "This paragraph contains ", new XElement("b", "bold"), " text."),
                    new XElement("p","This paragraph has just plain text.")
                 )
              )
            );
            return doc;
        }
        public string GetXMLString => GetXMLDoc().ToString();
        #endregion
        #region userdefaults class
        internal static class UserDefaults
        {
            //internal static string TitleString { get { return Properties.Settings.Default.defaultTitle; } set { Properties.Settings.Default.defaultTitle = value; } }
            //internal static string SubjectString { get { return Properties.Settings.Default.defaultSubject; } set { Properties.Settings.Default.defaultSubject = value; } }
            internal static string AuthorString { get { return Properties.Settings.Default.defaultAuthor; } set { Properties.Settings.Default.defaultAuthor = value; } }
            internal static string ManagerString { get { return Properties.Settings.Default.defaultManager; } set { Properties.Settings.Default.defaultManager = value; } }
            internal static string CompanyString { get { return Properties.Settings.Default.defaultCompany; } set { Properties.Settings.Default.defaultCompany = value; } }
            internal static string ConfindentialityString { get { return Properties.Settings.Default.defaultConfidentiality; } set { Properties.Settings.Default.defaultConfidentiality = value; } }
            internal static string LogoImageFileName { get { return Properties.Settings.Default.LogoImagePath; } set { Properties.Settings.Default.LogoImagePath = value; } }
            internal static void Save()
            {
                Properties.Settings.Default.Save();
            }
        }
        #endregion
    }

}
