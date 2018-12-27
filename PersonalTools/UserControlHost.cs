using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace peteli.PersonalTools
{
    // to display WPF (Windows Presentation Foundation) in Office Custom Task Pane
    // it needs to be embedded in win32 form control -> ElementHostControl
    [ComVisible(true)] //this is important excel throws expception otherwise!
    public partial class UserControlHost: UserControl
    {   
        public UserControlHost()
        {
            InitializeComponent();
        }
    }
}
