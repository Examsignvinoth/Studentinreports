using Microsoft.Reporting.WinForms;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
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

namespace Studentinreports
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        //commplet


        SqlConnection con = new SqlConnection(Properties.Settings.Default.Connection);
        SqlCommand cmd = new SqlCommand();


        //Assingent values 
        string data = "";
        DateTime fromdate = new DateTime(2020, 07, 13);
        DateTime Todate = new DateTime(2020, 07, 18);

        string fromdate1;
        string Todate1;



        List<StudentDetails> StudentInformation = new List<StudentDetails>();
        List<ClassStudent> studentsno = new List<ClassStudent>();

        List<TotalClassAndAttendClassDuration> ClassesAttenceStatus = new List<TotalClassAndAttendClassDuration>();
        List<StudentAssingmentDetails> StudentAssingentStaus = new List<StudentAssingmentDetails>();

        List<FinalReport> StudentFinalReports = new List<FinalReport>();
        //   StudentDetails a = new StudentDetails();
        public MainWindow()
        {
            InitializeComponent();





        }


        public List<StudentDetails> GetStudentDetails(string PPSno)
        {
            StudentInformation.Clear();
            con.Open();
            cmd.Connection = con;
            cmd.CommandText = "select AD.Subject,AD.CourseId,sd.PPSNo,SD.UserId,sd.Name,sd.Sclass,SD.Section FROM StudentDetails as SD JOIN " +
                "ALLLanguageDetails as AD on SD.GoogleMeetId=AD.MailId where SD.PPSNo='" + data + "'";
            SqlDataReader reader1;
            reader1 = cmd.ExecuteReader();
            while (reader1.Read())
            {
                StudentInformation.Add(new StudentDetails
                {
                    Subjects = reader1.GetString(0),
                    Couseid = reader1.GetString(1),
                    PSSNO = reader1.GetString(2),
                    Userid = reader1.GetString(3),
                    Name = reader1.GetString(4),
                    Sclass = reader1.GetString(5),
                    Section = reader1.GetString(6)

                });

            }

            reader1.Close();
            con.Close();
            return StudentInformation;
        }
        public List<TotalClassAndAttendClassDuration> GetAttenceDetails(string PPSNO, DateTime from, DateTime to)
        {
            //TotalClasses
            ClassesAttenceStatus.Clear();

            foreach (var i in StudentInformation)
            {
                TotalClassAndAttendClassDuration ob = new TotalClassAndAttendClassDuration();

                using (SqlConnection con = new SqlConnection(Properties.Settings.Default.Connection))
                {


                    SqlCommand cmd1 = new SqlCommand();
                    cmd1.Connection = con;
                    con.Open();
                    cmd1.CommandText = " select Subject, count(distinct Meetingcode) as TotalClass from StudentAttendanceLogsNew where Sclass = '" + i.Sclass + "' and Subject='" + i.Subjects + "' and EventDate between '" + fromdate.ToString("yyyy-MM-dd") + "' and '" + Todate.ToString("yyyy-MM-dd") + "'  group by Subject";
                    SqlDataReader reader2;
                    reader2 = cmd1.ExecuteReader();
                    while (reader2.Read())
                    {
                        ob.Subject = reader2.GetString(0);
                        ob.TotalClasses = reader2.GetInt32(1);

                    }
                    reader2.Close();
                    con.Close();

                }

                using (SqlConnection con = new SqlConnection(Properties.Settings.Default.Connection))
                {

                    SqlCommand cmd2 = new SqlCommand();
                    cmd2.Connection = con;
                    con.Open();
                    cmd2.CommandText = "select distinct TimeTableNew.Duration from StudentDetails JOIN TimeTableNew on StudentDetails.Sclass = TimeTableNew.Grade where Sclass='" + i.Sclass + "' and Subject='" + i.Subjects + "'";
                    SqlDataReader reader3;
                    reader3 = cmd2.ExecuteReader();
                    while (reader3.Read())
                    {
                        ob.ActualTotalDuration = reader3.GetString(0);
                    }
                    reader3.Close();
                    con.Close();
                }

                using (SqlConnection con3 = new SqlConnection(Properties.Settings.Default.Connection))
                {



                    SqlCommand cmd3 = new SqlCommand();
                    cmd3.Connection = con;
                    con.Open();
                    cmd3.CommandText = "select count ( SA.Subject) as AttendedClasses,sum(SA.TotalDuration) as AttendedClassesDuration from StudentDetails as SD join StudentAttendanceLogsNew as SA on SD.GoogleMeetId=SA.ParticipantId where SD.PPSNo='" + data + "' and SA.Sclass='" + i.Sclass + "' and sa.Subject='" + i.Subjects + "' and SA.EventDate between  '" + fromdate.ToString("yyyy-MM-dd") + "' and '" + Todate.ToString("yyyy-MM-dd") + "' group by SA.Subject ";
                    SqlDataReader reader4;
                    reader4 = cmd3.ExecuteReader();
                    while (reader4.Read())
                    {
                        ob.AttendedClasses = reader4.GetInt32(0);
                        ob.AttendedClassesDuration = reader4.GetInt32(1);
                    }
                    reader4.Close();
                    con.Close();
                }

                ob.TotalClassesDuration = ob.AttendedClasses * int.Parse(ob.ActualTotalDuration);

                if (ob.AttendedClassesDuration < ob.TotalClassesDuration)
                {

                    ob.TotalAttendClassesDuration = Math.Round(ob.AttendedClassesDuration * 1.0 / ob.TotalClassesDuration * 100);
                }
                else
                {
                    ob.TotalAttendClassesDuration = Math.Round(ob.TotalClassesDuration * 1.0 / ob.TotalClassesDuration * 100);
                }

                ob.AttendPercentage = Math.Round((ob.AttendedClasses*1.0 / ob.TotalClasses * 100),1);

                if (double.IsNaN(ob.TotalAttendClassesDuration))
                {
                    ob.TotalAttendClassesDuration = 0;
                }



                if (ob.TotalClasses != 0)
                {
                    ClassesAttenceStatus.Add(ob);
                }


            }



            return ClassesAttenceStatus;

        }

        public List<StudentAssingmentDetails> GetAssingmetnDetails(string PPSNO, DateTime from, DateTime to)
        {
            StudentAssingentStaus.Clear();


            foreach (var i in StudentInformation)
            {
                StudentAssingmentDetails ob = new StudentAssingmentDetails();

                using (SqlConnection con = new SqlConnection(Properties.Settings.Default.Connection))
                {

                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = con;
                    con.Open();
                    cmd.CommandText = "select  distinct CourseId, count(CreationTime) as TotalAssingments,sum(cast (MaxPoint as int)) as MaxPointAssingment from CourseWorklog  where CourseId in(select distinct a1.CourseId from  ALLLanguageDetails a1 join StudentDetails s1 on a1.UserId= s1.UserId and s1.PPSNo='" + data + "') and CourseId='" + i.Couseid + "' and CreationTime between '" + fromdate.ToString("yyyy-MM-dd") + "' and '" + Todate.ToString("yyyy-MM-dd") + "' group by CourseId";
                    SqlDataReader reader1;
                    reader1 = cmd.ExecuteReader();
                    while (reader1.Read())
                    {
                        ob.CourseId = reader1.GetString(0);
                        ob.TotalAssingments = reader1.GetInt32(1);
                        ob.MaxPointAssingment = reader1.GetInt32(2);
                    }
                    reader1.Close();
                    con.Close();

                }

                using (SqlConnection con = new SqlConnection(Properties.Settings.Default.Connection))
                {

                    double dvalue = 0.0;
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = con;
                    con.Open();
                    cmd.CommandText = " select s1.CourseId,count(s1.CourseWorkId) as OnTimesubmission,s1.AssignedGrade as OnTimeMarks from Submissionlog s1 join CourseWorklog c1 on s1.CourseWorkId = c1.CourseWorkId where s1.UserId = '" + i.Userid + "' and s1.CourseId = '" + i.Couseid + "' and c1.MaxPoint != 0 and s1.State in('RETURNED', 'TURNED_IN', 'RECLAIMED_BY_STUDENT')and IsLate = 0 and c1.CreationTime between '" + fromdate.ToString("yyyy-MM-dd") + "' and '" + Todate.ToString("yyyy-MM-dd") + "' group by  s1.CourseId ,s1.AssignedGrade,s1.CourseWorkId";
                    SqlDataReader reader2;
                    reader2 = cmd.ExecuteReader();
                    while (reader2.Read())
                    {
                        ob.OnTimesubmission += reader2.GetInt32(1);
                        if (string.IsNullOrEmpty(reader2.GetString(2)))
                        {
                            dvalue = 0;
                        }
                        else
                        {
                            dvalue = Convert.ToDouble(reader2.GetString(2));
                        }
                        ob.OnTimeMarks = ob.OnTimeMarks + dvalue;
                    }
                    reader2.Close();
                    con.Close();

                }
                using (SqlConnection con = new SqlConnection(Properties.Settings.Default.Connection))
                {

                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = con;
                    con.Open();
                    cmd.CommandText = "select s1.CourseId,COUNT(s1.CourseWorkId) Latesubmission, s1.AssignedGrade   as LateTimeMarks from Submissionlog s1 join CourseWorklog c1 on s1.CourseWorkId = c1.CourseWorkId where s1.UserId = '" + i.Userid + "' and s1.CourseId ='" + i.Couseid + "' and c1.MaxPoint != 0  and s1.State in('RETURNED', 'TURNED_IN', 'RECLAIMED_BY_STUDENT') and IsLate = 1 and c1.CreationTime between '" + fromdate.ToString("yyyy-MM-dd") + "' and '" + Todate.ToString("yyyy-MM-dd") + "' group by  s1.CourseId,s1.CourseWorkId ,s1.AssignedGrade";
                    SqlDataReader reader3;
                    reader3 = cmd.ExecuteReader();
                    while (reader3.Read())
                    {

                        ob.Latesubmission += reader3.GetInt32(1);
                        if (string.IsNullOrEmpty(reader3.GetString(2)))
                        {
                            ob.LateTimeMarks += 0;
                        }
                        else
                        {
                            ob.LateTimeMarks += Convert.ToDecimal(float.Parse(reader3.GetString(2)));

                        }
                        //  ob.Latesubmission = reader3.GetInt32(1);
                        // ob.LateTimeMarks = reader3.GetInt32(2);
                    }
                    reader3.Close();
                    con.Close();

                }

                using (SqlConnection con = new SqlConnection(Properties.Settings.Default.Connection))
                {

                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = con;
                    con.Open();
                    cmd.CommandText = "select  s1.CourseId,count(s1.CourseWorkId) Notmission from Submissionlog s1 join CourseWorklog c1 on s1.CourseWorkId=c1.CourseWorkId where s1.UserId='" + i.Userid + "' and s1.CourseId='" + i.Couseid + "' and c1.MaxPoint!=0  and s1.State in('NEW','CREATED') and IsLate=1 and c1.CreationTime between '" + fromdate.ToString("yyyy-MM-dd") + "' and '" + Todate.ToString("yyyy-MM-dd") + "' group by  s1.CourseId";
                    SqlDataReader reader4;
                    reader4 = cmd.ExecuteReader();
                    while (reader4.Read())
                    {
                        ob.Notmission = reader4.GetInt32(1);

                    }
                    reader4.Close();
                    con.Close();

                }



                Double temp = Convert.ToDouble(ob.OnTimeMarks) + Convert.ToDouble(ob.LateTimeMarks);
                Double temp1 = ob.MaxPointAssingment;
                Double temp3 = Math.Round((temp / temp1 * 10), 1);
                ob.TotalAvgMarks = temp3;

                if (double.IsNaN(ob.TotalAvgMarks))
                {
                    ob.TotalAvgMarks = 0;
                }


                if (ob.TotalAssingments != 0)
                {

                    StudentAssingentStaus.Add(ob);


                }


            }


            if (StudentAssingentStaus.Count != 0)

            {
                for (int j = 0; j < StudentInformation.Count(); j++)
                {
                    if (StudentInformation[j].Couseid == StudentAssingentStaus[j].CourseId)
                    {

                        StudentAssingentStaus[j].Subject = StudentInformation[j].Subjects;


                    }

                }

            }
            // System.Windows.Forms.MessageBox.Show("tecfe");

            return StudentAssingentStaus;
        }

        private void ATTENANCE_Click(object sender, RoutedEventArgs e)
        {
            fromdate1 = Convert.ToString(fromdate.ToShortDateString());
            Todate1 = Convert.ToString(Todate.ToShortDateString());



            for (int x = 6; x <= 6; x++)
            {
                studentsno.Clear();
                using (SqlConnection con = new SqlConnection(Properties.Settings.Default.Connection))
                {
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = con;
                    con.Open();

                    cmd.CommandText = "  select top 5 PPSNo from StudentDetails where Sclass=" + x + " order by PPSNo asc";

                    SqlDataReader reader1;


                    reader1 = cmd.ExecuteReader();
                    while (reader1.Read())
                    {
                        studentsno.Add(new ClassStudent
                        {
                            PPSNO = reader1.GetString(0)

                        });

                    }


                    reader1.Close();


                }
                foreach (var a in studentsno)

                {
                    data = a.PPSNO;
                    GetStudentDetails(data);
                    GetAttenceDetails(data, fromdate, Todate);
                    GetAssingmetnDetails(data, fromdate, Todate);
                    StudentFinalReports.Clear();

                    foreach (var z in StudentInformation)
                    {
                        StudentFinalReports.Add(new FinalReport
                        {
                            PPsno = z.PSSNO,

                            StudentName = z.Name,

                            ClassNo = z.Sclass,
                            fromDate = Convert.ToDateTime(fromdate1).ToString("dd-MM-yyyy"),
                            ToDate = Convert.ToDateTime(Todate1).ToString("dd-MM-yyyy"),
                            // = System.DateTime.Now.ToString(),
                            Section1 = z.Section
                        });
                    }


                    foreach (var i in ClassesAttenceStatus)
                    {



                        StudentFinalReports.Add(new FinalReport
                        {
                            Subjects = i.Subject.ToUpper(),
                            TotalClasses = i.TotalClasses,
                            AttendedClasses = i.AttendedClasses,
                            TotalAttendClassesDuration = i.TotalAttendClassesDuration,
                            AttendedPercentages = i.AttendPercentage,
                            // TotalAttendClassesDuration = attendduration,
                            Status = "Attence",
                            //PPsno = StudentInformation[0].PSSNO,

                            //StudentName = StudentInformation[0].Name,

                            //ClassNo = StudentInformation[0].Sclass,
                            //fromDate = Convert.ToDateTime(fromdate1).ToString("dd-MM-yyyy"),
                            //ToDate = Convert.ToDateTime(Todate1).ToString("dd-MM-yyyy"),
                            //// = System.DateTime.Now.ToString(),
                            //Section1 = StudentInformation[0].Section





                        });

                    }
                    foreach (var j in StudentAssingentStaus)
                    {
                        StudentFinalReports.Add(new FinalReport
                        {
                            Subjectes = j.Subject,
                            TotalAssingments = j.TotalAssingments,
                            OnTimesubmission = j.OnTimesubmission,
                            Latesubmission = j.Latesubmission,
                            Notmission = j.Notmission,
                            TotalAvgMarks = j.TotalAvgMarks,
                            Status = "Assingment"

                        });
                    }

                    if (StudentFinalReports.LongCount() != 0)
                    {
                        FillData3(StudentFinalReports);
                    }
                    else
                    {
                        System.Windows.Forms.MessageBox.Show("No Record" + fromdate1 + "  to  " + Todate1);
                        break;
                    }



                }



            }



            //GetAttenceDetails(data, fromdate, Todate);
            //GetAssingmetnDetails(data, fromdate, Todate);





            // FillData3(StudentFinalReports);

        }


        public class ClassStudent
        {
            public string PPSNO { get; set; }
        }
        public class StudentDetails
        {
            public string Subjects { get; set; }

            public string Couseid { get; set; }

            public string PSSNO { get; set; }


            public string Userid { get; set; }
            public string Name { get; set; }

            public string Sclass { get; set; }
            public string Section { get; set; }





        }

        public class TotalClassAndAttendClassDuration
        {
            public string Subject { get; set; }
            public int TotalClasses { get; set; }
            public string ActualTotalDuration { get; set; }

            public int AttendedClasses { get; set; }

            public int AttendedClassesDuration { get; set; }

            public int TotalClassesDuration { get; set; }

            public Double TotalAttendClassesDuration { get; set; }

            public Double AttendPercentage { get; set; }

        }

        public class StudentAssingmentDetails
        {

            public string Subject { get; set; }
            public string CourseId { get; set; }
            public int TotalAssingments { get; set; }

            public int MaxPointAssingment { get; set; }

            public int OnTimesubmission { get; set; }

            public double OnTimeMarks { get; set; }

            public int Latesubmission { get; set; }

            public decimal LateTimeMarks { get; set; }
            public int Notmission { get; set; }

            public Double TotalAvgMarks { get; set; }
        }

        public class FinalReport
        {
            public string Subjects { get; set; }
            public int TotalClasses { get; set; }

            public int AttendedClasses { get; set; }
            public Double TotalAttendClassesDuration { get; set; }

            //  public Double AttendedPercentages { get; set; }

            public string Subjectes { get; set; }
            public int TotalAssingments { get; set; }

            public int OnTimesubmission { get; set; }

            public int Latesubmission { get; set; }

            public int Notmission { get; set; }

            public Double TotalAvgMarks { get; set; }

            public string Status { get; set; }
            public string PPsno { get; set; }
            public string ClassNo { get; set; }
            public string fromDate { get; set; }
            public string ToDate { get; set; }

            public Double AttendedPercentages { get; set; }
            public string StudentName { get; set; }
            public string Section1 { get; set; }






        }


        void FillData3(List<FinalReport> studentdetail)
        {
            DataTable dt1 = new DataTable();
            dt1 = ConvertToDataTable(studentdetail.ToList());
            ReportDataSource reportDataSource1 = new ReportDataSource();
            reportDataSource1.Name = "DataSet1"; // Name of the DataSet we set in .rdlc
            reportDataSource1.Value = dt1;
            ReportViewer reportViewer1 = new ReportViewer();
            reportViewer1.LocalReport.ReportEmbeddedResource = "Studentinreports.Report4.rdlc";
            reportViewer1.LocalReport.DataSources.Add(reportDataSource1);
            reportViewer1.RefreshReport();
            reportViewer1.ProcessingMode = ProcessingMode.Local;

            Warning[] warnings1;
            string[] streamids1;
            string mimeType1;
            string encoding1;
            string extension1;
            try
            {


                byte[] bytes = reportViewer1.LocalReport.Render(
                  "PDF", null, out mimeType1, out encoding1, out extension1,
                  out streamids1, out warnings1);
                string dir = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "OMRROOT\\" + "22" + "\\Report\\PrintReport\\OverAll\\");
                if (!Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }
                if (Directory.Exists(dir))
                {
                    var temp = "" + data + "" + "_Report";

                    FileStream fs = new FileStream(dir + temp + ".pdf", FileMode.Create);

                    var temps = fs.ToString();
                    fs.Write(bytes, 0, bytes.Length);
                    fs.Close();
                }
                string path = dir + "" + data + "" + "_Report" + ".pdf";
                FileInfo fi = new FileInfo(path);
                if (fi.Exists)
                {
                    System.Diagnostics.Process.Start(path);
                }
                else
                {
                    System.Windows.MessageBox.Show("File not found");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


        public DataTable ConvertToDataTable<T>(IList<T> data)
        {
            PropertyDescriptorCollection properties =
               TypeDescriptor.GetProperties(typeof(T));
            DataTable table = new DataTable();
            foreach (PropertyDescriptor prop in properties)
                table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
            foreach (T item in data)
            {
                DataRow row = table.NewRow();
                foreach (PropertyDescriptor prop in properties)
                    row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;
                table.Rows.Add(row);
            }
            return table;

        }





    }
}
