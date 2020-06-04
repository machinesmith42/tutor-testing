using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Interop.Access;
using System.Data;
using DataTable = System.Data.DataTable;
using System.Data.Common;
using tutor_testing_v3.TutorDataSetTableAdapters;
using System.Diagnostics;

namespace tutor_testing_v3{
    class Program{
        public static Microsoft.Office.Interop.PowerPoint.Application application = new Microsoft.Office.Interop.PowerPoint.Application();
        public static Presentations ppPresens = application.Presentations;
        public static Presentation objPres = ppPresens.Open(AppDomain.CurrentDomain.BaseDirectory + "\\assets\\better powerpoint test v2.pptm", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoTrue);
        public static Slides objSlides = objPres.Slides;
        public static SlideShowSettings objSSS = objPres.SlideShowSettings;
        
        public static TutorDataSet.AllTutorsDataTable tutorTable = new TutorDataSet.AllTutorsDataTable();
        public static TutorDataSet.ScheduleDataTable scheduleTable = new TutorDataSet.ScheduleDataTable();
        public static TutorDataSet.SubjectDataTable subjectTable = new TutorDataSet.SubjectDataTable();

        static void Main(string[] args) {
            Init();
            MainLoop();
        }
        static void Init() {
            TutorDataSet db = new TutorDataSet();
            db.Clear();
            application.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
            objSSS.Run();
            AllTutorsTableAdapter tutorTableAdapt = new AllTutorsTableAdapter();
            tutorTableAdapt.Fill(tutorTable);
            ScheduleTableAdapter scheduleAdapt = new ScheduleTableAdapter();
            scheduleAdapt.Fill(scheduleTable);
            SubjectTableAdapter subjectAdapt = new SubjectTableAdapter();
            subjectAdapt.Fill(subjectTable);

        }
        static void MainLoop() {
            //while (true) {
            DisplayTutors();
            //}
        }
        static void DisplayTutors() {
            int currentDay = (int)DateTime.Now.DayOfWeek;
            var query =
                from tutor in tutorTable.AsEnumerable()
                join schedule in scheduleTable
                on tutor.Field<int>("ID") equals schedule.Field<int>("ID")
                where schedule.Field<int>("Day") == currentDay
                select new {
                    TutorID = tutor.Field <int> ("ID"),
                    Name = tutor.Field <string> ("FirstName") + " " + tutor.Field <string> ("LastName")
                };

            foreach(var q in query) {
                string printLine = q.Name + " " + q.TutorID.ToString();
            }
        }
    }
}
