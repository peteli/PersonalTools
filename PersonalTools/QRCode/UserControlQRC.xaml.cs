using System;
using System.Collections.Generic;
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
using QRCoder;
using System.Drawing;
using System.Windows.Interop;

namespace peteli.PersonalTools.QR
{
    /// <summary>
    /// Interaktionslogik für UserControlQRC.xaml
    /// </summary>
    public partial class UserControlQRC : UserControl
    {
        public UserControlQRC()
        {
            InitializeComponent();
            this.btnMakeQRCode.Click += BtnMakeQRCode_Click;
        }

        private void BtnMakeQRCode_Click(object sender, RoutedEventArgs e)
        {
            string QRinput = this.QRcontent.Text;
            if (string.IsNullOrWhiteSpace(QRinput))
            {
                QRinput = "Need Input, Dude!";
             };
            QRCodeGenerator qrGenerator = new QRCodeGenerator();
            QRCodeData qrCodeData = qrGenerator.CreateQrCode(QRinput, QRCodeGenerator.ECCLevel.Q);
            QRCoder.QRCode qrCode = new QRCoder.QRCode(qrCodeData);
            Bitmap qrCodeImage = qrCode.GetGraphic(20);
            this.ImageQRCode.Source = Imaging.CreateBitmapSourceFromHBitmap(qrCodeImage.GetHbitmap(), IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions()); ;
        }
    }
}
