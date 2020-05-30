using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Interop.Access;

namespace tutor_testing_v3
{
    class Program
    {
        public static OleDbConnection myConn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "\\assets\\Tutor.accdb");
        public static Microsoft.Office.Interop.PowerPoint.Application application = new Microsoft.Office.Interop.PowerPoint.Application();
        public static Presentations ppPresens = application.Presentations;
        public static Presentation objPres = ppPresens.Open(AppDomain.CurrentDomain.BaseDirectory + "\\assets\\better powerpoint test v2.pptm", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoTrue);
        public static Slides objSlides = objPres.Slides;
        public static SlideShowSettings objSSS = objPres.SlideShowSettings;
        static void Main(string[] args)
        {
            Init();
        }
        static void Init() {
            myConn.Open();
            application.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
            objSSS.Run();
        }
    }
}
