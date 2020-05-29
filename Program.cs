using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Net.Security;
using System.Text.Encodings.Web;
using System.Web.Script.Serialization;
using WebSocketSharp;
using MySql.Data.MySqlClient;
using System.Configuration;
using System.Net.Mail;
using System.Data;

namespace mailautomation
{
       
    class Program
    {
        static  MySqlConnection conn = null;
        static MySqlConnection intconn = null;
        static MySqlCommand cmd = null;
        static MySqlDataReader rdr = null;
        static MySqlDataAdapter adapter = null;
        private static string connectionString1 = ConfigurationManager.ConnectionStrings["cuea"].ConnectionString;
        private static string connectionString = ConfigurationManager.ConnectionStrings["cuea_ac"].ConnectionString;
        private static string connectionString2 = ConfigurationManager.ConnectionStrings["cuea_tna"].ConnectionString;
        static MySqlCommandBuilder cmdB = null;
        static DataTable dt = null;
       static void Main(string[] args)
        {
            try
            {
                var result = Connectwsa();
                Console.WriteLine("Results" + result);
                System.Threading.Thread.Sleep(300000);
                SendSupervisor_email_test();
                Console.WriteLine("Emails Sents all");
                //Console.ReadLine();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }

            
       }
       
        public static async Task Connectwsa()
        {
            string messg = string.Empty;
			string messgr = string.Empty;
            DateTime now = DateTime.Today;
            messg = "{\"event" + "\":\"" + "REPORT" + "\",\"" + "data" + "\":\"" + "{" + "\\" + "\"limit" + "\\" + "\":50," + "\\" + "\"offset" + "\\" + "\":0," + "\\" + "\"type" + "\\" + "\":" + "\\" + "\"callCreateReport" + "\\" + "\"," + "\\" + "\"start_datetime" + "\\" + "\":" + "\\" + "\"" + now.ToString("yyyy-MM-dd") + "\\" + "\"," + "\\" + "\"end_datetime" + "\\" + "\":" + "\\" + "\"" + now.ToString("yyyy-MM-dd") + "\\" + "\"," + "\\" + "\"report_type" + "\\" + "\":" + "\\" + "\"REPORT_DAILY" + "\\" + "\"," + "\\" + "\"group_id_list" + "\\" + "\":[4587,8242,8243,8251]," + "\\" + "\"language" + "\\" + "\":" + "\\" + "\"en" + "\\" + "\"," + "\\" + "\"rebuild_time_card" + "\\" + "\":true," + "\\" + "\"columns" + "\\" + "\":[{" + "\\" + "\"name" + "\\" + "\":" + "\\" + "\"report.date" + "\\" + "\"," + "\\" + "\"field" + "\\" + "\":" + "\\" + "\"datetime" + "\\" + "\"," + "\\" + "\"displayName" + "\\" + "\":" + "\\" + "\"Date" + "\\" + "\"},{" + "\\" + "\"name" + "\\" + "\":" + "\\" + "\"report.userName" + "\\" + "\"," + "\\" + "\"field" + "\\" + "\":" + "\\" + "\"userName" + "\\" + "\"," + "\\" + "\"displayName" + "\\" + "\":" + "\\" + "\"Name" + "\\" + "\"},{" + "\\" + "\"name" + "\\" + "\":" + "\\" + "\"report.userId" + "\\" + "\"," + "\\" + "\"field" + "\\" + "\":" + "\\" + "\"userId" + "\\" + "\"," + "\\" + "\"displayName" + "\\" + "\":" + "\\" + "\"User ID" + "\\" + "\"},{" + "\\" + "\"name" + "\\" + "\":" + "\\" + "\"report.department" + "\\" + "\"," + "\\" + "\"field" + "\\" + "\":" + "\\" + "\"userGroupName" + "\\" + "\"," + "\\" + "\"displayName" + "\\" + "\":" + "\\" + "\"Department" + "\\" + "\"},{" + "\\" + "\"name" + "\\" + "\":" + "\\" + "\"report.shift" + "\\" + "\"," + "\\" + "\"field" + "\\" + "\":" + "\\" + "\"shift" + "\\" + "\"," + "\\" + "\"displayName" + "\\" + "\":" + "\\" + "\"Shift" + "\\" + "\"},{" + "\\" + "\"name" + "\\" + "\":" + "\\" + "\"report.leave" + "\\" + "\"," + "\\" + "\"field" + "\\" + "\":" + "\\" + "\"leave" + "\\" + "\"," + "\\" + "\"displayName" + "\\" + "\":" + "\\" + "\"Leave" + "\\" + "\"},{" + "\\" + "\"name" + "\\" + "\":" + "\\" + "\"timeCard.in" + "\\" + "\"," + "\\" + "\"field" + "\\" + "\":" + "\\" + "\"inTime" + "\\" + "\"," + "\\" + "\"displayName" + "\\" + "\":" + "\\" + "\"In" + "\\" + "\"},{" + "\\" + "\"name" + "\\" + "\":" + "\\" + "\"timeCard.out" + "\\" + "\"," + "\\" + "\"field" + "\\" + "\":" + "\\" + "\"outTime" + "\\" + "\"," + "\\" + "\"displayName" + "\\" + "\":" + "\\" + "\"Out" + "\\" + "\"},{" + "\\" + "\"name" + "\\" + "\":" + "\\" + "\"report.exception" + "\\" + "\"," + "\\" + "\"field" + "\\" + "\":" + "\\" + "\"exceptionForView" + "\\" + "\"," + "\\" + "\"displayName" + "\\" + "\":" + "\\" + "\"Exception" + "\\" + "\"},{" + "\\" + "\"name" + "\\" + "\":" + "\\" + "\"timeCard.regularHours" + "\\" + "\"," + "\\" + "\"field" + "\\" + "\":" + "\\" + "\"normalRegular" + "\\" + "\"," + "\\" + "\"displayName" + "\\" + "\":" + "\\" + "\"Regular hours" + "\\" + "\"},{" + "\\" + "\"name" + "\\" + "\":" + "\\" + "\"timeCard.overtimeHours" + "\\" + "\"," + "\\" + "\"field" + "\\" + "\":" + "\\" + "\"normalOvertime" + "\\" + "\"," + "\\" + "\"displayName" + "\\" + "\":" + "\\" + "\"Overtime hours" + "\\" + "\"},{" + "\\" + "\"name" + "\\" + "\":" + "\\" + "\"report.totalWorkTime" + "\\" + "\"," + "\\" + "\"field" + "\\" + "\":" + "\\" + "\"totalWorkTime" + "\\" + "\"," + "\\" + "\"displayName" + "\\" + "\":" + "\\" + "\"Total Work Hours" + "\\" + "\"}]}\"}";
            Console.WriteLine(messg);
            //{"event":"REPORT","data":{"limit":"50","offset":"0","type":"callCreateReport","start_datetime":"2020-05-01","end_datetime":"2020-05-31","report_type":"REPORT_DAILY","group_id_list":[8243],"language":"en","rebuild_time_card":true}"}
           
            var ws = new WebSocket("wss://127.0.0.1:3002/ws?language=en");
            
            //var ws = new WebSocket("wss://127.0.0.1/wsapi");
            ws.Connect();
            ws.OnMessage += (sender, e) =>
            {
                Console.WriteLine(e.Data);
				messgr= e.Data;
                //Console.WriteLine(e.RawData);
                };
            ws.Send(messg);                   
            ws.Close();
            await Task.Delay(300000);
            
        }

        public static void SendSupervisor_email_test()
        {
            try
            {
                int TotalCount = 0;
                DateTime now = DateTime.Today;
                intconn = new MySqlConnection(connectionString2 + "default command timeout=3600 ");
                intconn.Open();
                dt = new DataTable();
                string SQLQuery = "select distinct SupervisorID,ebt.email as email from supervisor,biostar_tna.user ebt where ebt.user_id=SupervisorID and SupervisorID=1003";
                cmd = new MySqlCommand(SQLQuery, intconn);
                MySqlDataAdapter adapter = new MySqlDataAdapter(cmd);
                adapter.Fill(dt);
                TotalCount = dt.Rows.Count;
                foreach (DataRow row in dt.Rows)
                {
                    DataTable dtEmployee = GetData(now.ToString("yyyy-MM-dd"), row["SupervisorID"].ToString());
                    CreateCSV(dtEmployee, ".");
                    sendmail(row["email"].ToString());
                    //this.sendmail(row["SupervisorID"].ToString(), row["email"].ToString());
                   
                }
                intconn.Close();
                intconn.Dispose();
            }
            catch (Exception err)
            {
                Console.WriteLine(err.Message);
            }
        }

        public static DataTable GetData(String st, String st3)
        {
            string conString = ConfigurationManager.ConnectionStrings["cuea_tna"].ConnectionString; ;
            MySqlConnection conn = new MySqlConnection(conString);
            DataSet ds = new DataSet();
            try
            {
                conn.Open();
                String sql = "SELECT tcd.user_id,replace(uds.NAME,',',' ') as Name,uds.PHONE,uds.EMAIL,replace(uds.CAMPUS,',',' ') as CAMPUS,replace(uds.DESIGNATION,',',' ') as DESIGNATION ,replace(uds.DIVISION,',',' ') as DIVISION,replace(uds.DEPARTMENT,',',' ') as DEPARTMENT ,replace(uds.CATEGORY,',',' ') as CATEGORY ,replace(uds.FACULTY,',',' ') as FACULTY,tcd.date,time(DATE_ADD(tcd.punch_in, INTERVAL 3 HOUR)) as punch_in,Time(DATE_ADD(tcd.punch_out, INTERVAL 3 HOUR)) as punch_out,SEC_TO_TIME(tcd.total_working_time) as TotalWorkHours,SEC_TO_TIME(tcd.regular_schedule_time) as RegularHours,SEC_TO_TIME(tcd.normal_regular_working_time) as RegularWorkHours , (CASE WHEN tcd.exception_codes = 'TA_EXCEPTION_2' THEN 'Absence' WHEN tcd.exception_codes = 'TA_EXCEPTION_3' THEN 'LateIn' WHEN tcd.exception_codes = 'TA_EXCEPTION_4' THEN 'Early_Out' WHEN tcd.exception_codes = 'TA_EXCEPTION_6,TA_EXCEPTION_3' THEN 'MissingPunchOut_LateIn'  WHEN tcd.exception_codes = 'TA_EXCEPTION_3,TA_EXCEPTION_4' THEN 'LateIn Early Out' WHEN tcd.exception_codes = 'TA_EXCEPTION_6' THEN 'MissingPunchOutEarlyOut' ELSE '-' END ) as Exception, tcd.shift_name,replace(get_mysupervisor_name(tcd.user_id),',','') as Supervisor_Name FROM timecard tcd, user_detail uds WHERE tcd.user_id = uds.usrid COLLATE utf8_unicode_ci AND tcd.date = '" + st + "' and tcd.user_id in(select user_id from supervisor where SupervisorID='" + st3 + "')";
                MySqlCommand cmd = new MySqlCommand(sql, conn);
                MySqlDataAdapter da = new MySqlDataAdapter();
                da.SelectCommand = cmd;
                da.Fill(ds);
            }
            catch (Exception ex)
            {
               Console.WriteLine(ex.ToString(), "Exception");
            }
            finally
            {
                conn.Close();
                conn.Dispose();
            }
            return ds.Tables[0];
        }

        public static void SendSupervisor_email()
        {
            try
            {
                int TotalCount = 0;
                intconn = new MySqlConnection(connectionString2 + "default command timeout=3600 ");
                intconn.Open();
                DateTime now = DateTime.Today;
                dt = new DataTable();
                string SQLQuery = "select distinct SupervisorID,ebt.email as email from supervisor,biostar_tna.user ebt where ebt.user_id=SupervisorID AND ebt.email IS NOT null";
                cmd = new MySqlCommand(SQLQuery, intconn);
                MySqlDataAdapter adapter = new MySqlDataAdapter(cmd);
                adapter.Fill(dt);
                TotalCount = dt.Rows.Count;
                foreach (DataRow row in dt.Rows)
                {
                    DataTable dtEmployee = GetData(now.ToString("yyyy-MM-dd"), row["SupervisorID"].ToString());
                    CreateCSV(dtEmployee, ".");
                    sendmail(row["email"].ToString());
                                        
                }
                intconn.Close();
                intconn.Dispose();
                Console.WriteLine("Mail Send Successfully to Suppervisor");
            }
            catch (Exception err)
            {
                Console.WriteLine(err.Message);
            }
        }

        protected static void CreateCSV(DataTable dt, string filePath)
        {
            System.Data.DataView view = new System.Data.DataView(dt);
            // for summary dt = view.ToTable(false, "user_id,NAME,PHONE,EMAIL,CAMPUS,DESIGNATION,DIVISION,DEPARTMENT,CATEGORY,FACULTY,TotalWorkHours,RegularHours,RegularWorkHours ,Absence,LateIn,Early_Out,MissingPunchOut_LateIn,LateInEarlyOut,MissingPunchOut ,tcd.shift_name");
            dt = view.ToTable(false, "user_id", "NAME", "PHONE", "EMAIL", "CAMPUS", "DESIGNATION", "DIVISION", "DEPARTMENT", "FACULTY", "DATE", "punch_in", "punch_out", "TotalWorkHours", "RegularHours", "RegularWorkHours", "Exception", "shift_name", "Supervisor_Name");
            StringBuilder sb = new StringBuilder();
            string[] columnNames = dt.Columns.Cast<DataColumn>().Select(column => column.ColumnName).ToArray();
            sb.AppendLine(string.Join(",", columnNames));
            foreach (DataRow row in dt.Rows)
            {
                string[] fields = row.ItemArray.Select(field => field.ToString()).ToArray();
                sb.AppendLine(string.Join(",", fields));
            }
            //Passing StringBuilder to Create CSV
            var bytes = Encoding.GetEncoding("iso-8859-1").GetBytes(sb.ToString());
            MemoryStream stream = new MemoryStream(bytes);
            string filename = "Dailyattendance.csv";
            string fileLocation = filePath + filename;
            using (var fileStream = new FileStream(fileLocation, FileMode.Create, FileAccess.Write))
            {
                stream.CopyTo(fileStream);
            }
        }

        public static void sendmail(String email)
        {
            try
            {

                MailMessage mail = new MailMessage();
                SmtpClient SmtpServer = new SmtpClient("smtp.gmail.com");
                mail.From = new MailAddress("CUEA Time and Attendance System <biostar@cuea.edu>");
                mail.To.Add(email.Trim());
                mail.Subject = "Daily Attendance  report From Automated Email";
                mail.Body = "Daily Attendance  Report";
                System.Net.Mail.Attachment attachment;
                attachment = new System.Net.Mail.Attachment(".Dailyattendance.csv");
                mail.Attachments.Add(attachment);
                SmtpServer.Port = 587;
                SmtpServer.UseDefaultCredentials = false;
                SmtpServer.EnableSsl = true;
                SmtpServer.DeliveryMethod = SmtpDeliveryMethod.Network;
                SmtpServer.Credentials = new System.Net.NetworkCredential("biostar@cuea.edu", "Biom3tr*Cuea");
                //SmtpServer.Timeout = 20000;
                SmtpServer.Send(mail);
                //MessageBox.Show("Mail Send Successfully to Suppervisor","",MessageBoxButtons.OK,MessageBoxIcon.Information);
                mail.Dispose();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString(), "Mail error");
            }
        }
    }
}
