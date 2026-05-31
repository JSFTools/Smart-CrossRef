using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;

namespace SmartCrossRef
{
    public partial class CrossRefPaneHostControl : UserControl
    {
        private winCrossRefPane wpfControl;

        public CrossRefPaneHostControl()
        {
            InitializeComponent();

            // Instantiate your WPF Control
            wpfControl = new winCrossRefPane();

            // Assign the WPF control to the ElementHost
            elementHost1.Child = wpfControl;
        }

        // Add this public method to bridge the call safely
        public void RefreshContextualData()
        {
            if (wpfControl != null)
            {
                wpfControl.PopulateContextualItems(scope: ScanScope.NearbyOnly);
            }
        }
        public void UpdateTheme()
        {
            try
            {
                // Query Office's current background theme
                int languageId = (int)Globals.ThisAddIn.Application.LanguageSettings.get_LanguageID(Office.MsoAppLanguageID.msoLanguageIDInstall);

                // Read the explicit Office Theme registry key (0 = Colorful, 3 = Dark Gray, 4 = Black)
                int themeValue = (int)Microsoft.Win32.Registry.GetValue(
                    @"HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common", "UI Theme", 0);


                // Office Theme Values:
                // 3 = Dark Gray
                // 4 = Black
                // 5 = White
                // 6  = Use System Setting (Windows Light/Dark Mode)
                // 7 = Colorful
                // 2. Handle "Use system setting" fallback
                if (themeValue == 6)
                {
                    // 0 = Windows Dark Mode (Maps to Office Black), 1 = Windows Light Mode
                    int appsUseLightTheme = (int)Microsoft.Win32.Registry.GetValue(
                        @"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize",
                        "AppsUseLightTheme", 1);

                    themeValue = (appsUseLightTheme == 0) ? 4 : 0;
                }

                // 3. Define and apply the matching color profile
                System.Windows.Media.Color bg;
                System.Windows.Media.Color text;
                System.Windows.Media.Color border;

                switch (themeValue)
                {
                    case 4: // Office Black Theme
                        bg = System.Windows.Media.Color.FromRgb(32, 32, 32);       // #202020
                        text = System.Windows.Media.Color.FromRgb(241, 241, 241);  // #F1F1F1
                        border = System.Windows.Media.Color.FromRgb(69, 69, 69);   // #454545
                        break;

                    case 3: // Office Dark Gray Theme
                        bg = System.Windows.Media.Color.FromRgb(102, 102, 102);    // #666666 (Word's native pane gray)
                        text = System.Windows.Media.Color.FromRgb(255, 255, 255);  // #FFFFFF
                        border = System.Windows.Media.Color.FromRgb(140, 140, 140); // #8C8C8C
                        break;

                    case 0: // Office Colorful / White Themes
                    default:
                        bg = System.Windows.Media.Color.FromRgb(255, 255, 255);    // #FFFFFF
                        text = System.Windows.Media.Color.FromRgb(43, 43, 43);      // #2B2B2B
                        border = System.Windows.Media.Color.FromRgb(210, 210, 210); // #D2D2D2
                        break;
                }

                // Apply to WPF Controls
                wpfControl.ApplyOfficeTheme(bg, text, border);

                // Sync the WinForms host container background color 
                this.BackColor = Color.FromArgb(bg.R, bg.G, bg.B);
            }
            catch
            {
                // Safety fallback if registry permissions are restricted
                wpfControl.ApplyOfficeTheme(
                    System.Windows.Media.Colors.White,
                    System.Windows.Media.Colors.Black,
                    System.Windows.Media.Colors.Gray);
            }
        }
    }
}
