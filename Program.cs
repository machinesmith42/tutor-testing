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
using Microsoft.Office.Core;
using MsoTriState = Microsoft.Office.Core.MsoTriState;

namespace tutor_testing_v3{
    class Program{
        public static Microsoft.Office.Interop.PowerPoint.Application application = new Microsoft.Office.Interop.PowerPoint.Application();
        public static Presentations ppPresens = application.Presentations;
        public static Presentation objPres = ppPresens.Open(AppDomain.CurrentDomain.BaseDirectory + "\\assets\\better powerpoint test v2.pptm", MsoTriState.msoFalse, MsoTriState.msoTrue, MsoTriState.msoTrue);
        public static Slides objSlides = objPres.Slides;
        public static SlideShowSettings objSSS = objPres.SlideShowSettings;
        public static SlideShowWindow objSSW;
        public static TutorDataSet.AllTutorsDataTable tutorTable = new TutorDataSet.AllTutorsDataTable();
        public static TutorDataSet.ScheduleDataTable scheduleTable = new TutorDataSet.ScheduleDataTable();
        public static TutorDataSet.SubjectDataTable subjectTable = new TutorDataSet.SubjectDataTable();
        public static int tutorsSlide = 1;
        public static bool canDelete = false;

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
            objSSW = objPres.SlideShowWindow;
            

        }
        static void MainLoop() {
            while (true) {
                if (objSSW.View.Slide.SlideIndex == 4 && canDelete) {
                    DeleteSlides();
                    canDelete = false;
                } else if (objSSW.View.Slide.SlideIndex == 5) {
                    canDelete = true;
                    DisplayTutors();
                }
                
            }
        }
        internal static dynamic CurrentSlide {
            get {
                if (application.Active == MsoTriState.msoTrue &&
                    application.ActiveWindow.Panes[2].Active == MsoTriState.msoTrue) {
                    return application.ActiveWindow.View.Slide.SlideIndex;
                }
                return null;
            }
        }
        static void DisplayTutors() {
            DateTime currentDayTime = DateTime.Now;
            var query =
                from tutor in tutorTable.AsEnumerable()
                join schedule in scheduleTable
                on tutor.Field<int>("ID") equals schedule.Field<int>("ID")
                where schedule.Field<int>("Day") == (int)currentDayTime.DayOfWeek + 1&&
                schedule.Field<DateTime>("Start").TimeOfDay <= currentDayTime.TimeOfDay &&
                schedule.Field<DateTime>("End").TimeOfDay >= currentDayTime.TimeOfDay
                select new {
                    TutorID = tutor.Field <int> ("ID"),
                    Name = tutor.Field <string> ("FirstName") + " " + tutor.Field <string> ("LastName")
                };

            foreach(var q in query) {
                SlideRange slide = CreateSlide(tutorsSlide);
                WriteToTextbox(slide, "TutorName", q.Name);
            }
        }
        static SlideRange CreateSlide(int copyOfIndex) {
            SlideRange newSlide = objSlides[copyOfIndex].Duplicate();
            newSlide.SlideShowTransition.Hidden = MsoTriState.msoFalse;
            newSlide.Tags.Add("isCreated", "true");
            newSlide.MoveTo(objSlides.Count);
            return newSlide;
        }
        static string WriteToTextbox(SlideRange slide, string textboxName, string inputString) {
            slide.Shapes[textboxName].TextFrame.TextRange.Text = inputString;
            return inputString;
        }
        static int DeleteSlides() {
            int numberDeleted = 0;
            while(objSlides[objSlides.Count].Tags["isCreated"] == "true") {
                numberDeleted++;
                objSlides[objSlides.Count].Delete();
            }
            return numberDeleted;
        }
    }
};
