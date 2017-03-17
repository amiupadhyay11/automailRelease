using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Mail;
using System.Net;
using System.IO;
using System.Reflection;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace AutoMail
{
   

    class Program
    {
        string smtpAddress = "";
        int portNumber = 0;
        bool enableSSL = true;
        string emailFrom = "";
        string password = "";
        string emailTo = "";
        string subject = "";
        string bodyFinal = "";
        string bccEmail = "";
        string ccEmail = "";
        string MasterexcelSheetPath = "";
        string clientName1, clientName2, clientName3 = "";
        string year = DateTime.Now.Year.ToString();
        string monthnumber = DateTime.Now.ToString("yyyy-MM-MMM");
        string currentdate = DateTime.Now.ToString("yyyy-MM-dd");
        string fileName = "";
        string filePath = "";
        DataTable ContentTable = null;
        DataTable clientReportFile = null;
        DataTable ClientEmail = null;
        DataTable attachments = null;
   
      
        static void Main(string[] args)
        {
            Program p = new Program();
            p.ReadConfig();
            p.Readexcel();
            p.sendMail();
            
            

        }

        //Read config file
       
        public void ReadConfig()
        {
            //var data = File
            // .ReadAllLines(@"D:\Amit Programs\Automatic mail sending\AutoMail\Parameter\parameter.txt")
            // .Select(x => x.Split('='))
            // .Where(x => x.Length > 1)
            // .ToDictionary(x => x[0].Trim(), x => x[1]);

        //test json


            JObject o1 = JObject.Parse(File.ReadAllText(@"D:\Amit Programs\Automatic mail sending\AutoMail\Parameter\parameter.json"));


            // read JSON directly from a file
            using (StreamReader file = File.OpenText(@"D:\Amit Programs\Automatic mail sending\AutoMail\Parameter\parameter.json"))
            using (JsonTextReader reader = new JsonTextReader(file))
            {
                JObject o2 = (JObject)JToken.ReadFrom(reader);
                smtpAddress=o2["smtpAddress"].ToString().Trim();
                portNumber = Convert.ToInt32(o2["portName"].ToString().Trim());
                emailFrom=o2["emailFrom"].ToString().Trim();
                password =o2["password"].ToString().Trim();
                ccEmail =o2["ccEmail"].ToString().Trim();
                bccEmail =o2["bccEmail"].ToString().Trim();
                subject =o2["subject"].ToString().Trim();
                filePath = o2["pdfPath"].ToString().Trim() + "\\" + year + "\\" + monthnumber + "\\" + currentdate + "\\" + "report"; ;
                MasterexcelSheetPath=o2["MasterexcelSheetPath"].ToString().Trim();
                JArray body = (JArray)o2["body"];
                for (int i = 0; i < body.Count; i++)
                {
                    bodyFinal = bodyFinal + Environment.NewLine + body[i].ToString().Trim();
                    bodyFinal = bodyFinal.Replace(Environment.NewLine, "<br />");
                                                         
                    
                }
            

            }



          //  smtpAddress = data["smtpAddress"].ToString().Trim() ;
          //  portNumber =Convert.ToInt32(data["portName"].ToString().Trim());
          // emailFrom = data["emailFrom"].ToString().Trim();
          //  password = data["password"].ToString().Trim();
          //  bccEmail = data["bccEmail"].ToString().Trim();
          //  ccEmail = data["ccEmail"].ToString().Trim();
          ////  emailTo = data["emailTo"].ToString().Trim();
          //  subject = data["subject"].ToString().Trim();
          //  body = data["body"].ToString().Trim();
          //  filePath = data["pdfPath"].ToString().Trim()+"\\"+year+"\\"+monthnumber+"\\"+currentdate+"\\"+"report";
          //  MasterexcelSheetPath = data["MasterexcelSheetPath"].ToString().Trim();
          ////filePath = data["pdfPath"].ToString().Trim();
         
        }

        //read excel file

        public void Readexcel()
        {
            try
            {
               
                try
                {
                    string connString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="+MasterexcelSheetPath+";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=1\"";
                    DataTable schemaTable = new DataTable();
                    OleDbConnection OledbConn = new OleDbConnection(connString);   
                    OleDbCommand OledbCmd = new OleDbCommand();
                    OledbCmd.Connection = OledbConn;
                    OledbConn.Open();
                    OledbCmd.CommandText = "Select * from [Sheet1$]";
                    OleDbDataReader dr = OledbCmd.ExecuteReader();
                   // DataTable ContentTable = null;
                    if (dr.HasRows)
                    {
                        ContentTable = new DataTable();
                        ContentTable.Columns.Add("ClientID", typeof(string));
                        ContentTable.Columns.Add("ClientName1", typeof(string));
                        ContentTable.Columns.Add("ClientName2", typeof(string));
                        ContentTable.Columns.Add("ClientName3", typeof(string));
                        ContentTable.Columns.Add("Email1", typeof(string));
                        ContentTable.Columns.Add("Email2", typeof(string));
                        ContentTable.Columns.Add("Email3", typeof(string));
                        while (dr.Read())
                        {
                               // if (dr[0].ToString().Trim() != string.Empty && dr[1].ToString().Trim() != string.Empty && dr[2].ToString().Trim() != string.Empty && dr[0].ToString().Trim() != " " && dr[1].ToString().Trim() != " " && dr[2].ToString().Trim() != " ")
                            ContentTable.Rows.Add(dr[0].ToString().Trim(), dr[1].ToString().Trim(), dr[2].ToString().Trim(), dr[3].ToString().Trim(), dr[4].ToString().Trim(), dr[5].ToString().Trim(), dr[6].ToString().Trim());
                                

                        }
                    }
                    dr.Close();
                   
                    OledbConn.Close();
                    //return ContentTable;
                    
                    //Person table
                    OledbConn.Open();
                    OledbCmd.CommandText = "Select * from [Sheet2$]";
                    OleDbDataReader dr2 = OledbCmd.ExecuteReader();
                   
                    if (dr2.HasRows)
                    {
                        ClientEmail = new DataTable();
                        ClientEmail.Columns.Add("ClientName", typeof(string));
                        ClientEmail.Columns.Add("ClientEmail", typeof(string));
                        
                        while (dr2.Read())
                        {
                            // if (dr[0].ToString().Trim() != string.Empty && dr[1].ToString().Trim() != string.Empty && dr[2].ToString().Trim() != string.Empty && dr[0].ToString().Trim() != " " && dr[1].ToString().Trim() != " " && dr[2].ToString().Trim() != " ")
                            ClientEmail.Rows.Add(dr2[0].ToString().Trim(), dr2[1].ToString().Trim());
                           
                            
                        }
                    }
                    dr.Close();

                    OledbConn.Close();
                    
                }
                catch (Exception ex)
                {
                    throw ex;

                }    
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        //method to send mail
        public void sendMail()
        {

          
           
            LogWriter logwrite = new LogWriter();
           
                // Can set to false, if you are sending pure text.

            MessageBox.Show("Mail sending process started. Please press OK to proceed...");

                if (Directory.Exists(filePath))
                {
                    if (System.IO.Directory.GetFiles(filePath, "*", SearchOption.AllDirectories).Length != 0)
                    {
                        try
                        {
                            
                           int totcount = 0;
                              
                                        clientReportFile = new DataTable();
                                        clientReportFile.Columns.Add("ClientID", typeof(string));
                                        clientReportFile.Columns.Add("Filename", typeof(string));
                                        clientReportFile.Columns.Add("ClientName1", typeof(string));
                                        clientReportFile.Columns.Add("ClientName2", typeof(string));
                                        clientReportFile.Columns.Add("ClientName3", typeof(string));
                                        foreach (DataRow row in ContentTable.Rows)
                                        {
                                          string clientId = row["ClientID"].ToString().ToUpper();
                                          clientName1 = row["ClientName1"].ToString().ToUpper();
                                          clientName2 = row["ClientName2"].ToString().ToUpper();
                                          clientName3 = row["ClientName3"].ToString().ToUpper();

                                          foreach (FileInfo file in new DirectoryInfo(filePath).GetFiles("*.pdf"))
                                          {
                                              fileName = file.FullName.ToString();
                                              if (fileName.Contains(clientId))
                                              {
                                                  
                                                  
                                                  clientReportFile.Rows.Add(clientId.Trim(), fileName.Trim(),clientName1,clientName2,clientName3);
                                                  
                                                 
                                              }
                                          }

                                        
                                    }

                              
                               
                                fileName = "";

                                
                                foreach (DataRow row in ClientEmail.Rows)
                                {
                                    string clientName = row["ClientName"].ToString().ToUpper();
                                    emailTo = row["ClientEmail"].ToString();
                                    attachments = new DataTable();
                                    attachments.Columns.Add("AttachmentName", typeof(string));
                                    foreach (DataRow rowClientReport in clientReportFile.Rows)
                                    {
                                        
                                        string testname1 = rowClientReport["ClientName1"].ToString().ToUpper();
                                        string testname2 = rowClientReport["ClientName2"].ToString().ToUpper();
                                        string testname3 = rowClientReport["ClientName3"].ToString().ToUpper();
                                        //attachments = new DataTable();
                                        //attachments.Columns.Add("AttachmentName", typeof(string));
                                        if (clientName == rowClientReport["ClientName1"].ToString().ToUpper() || clientName == rowClientReport["ClientName2"].ToString().ToUpper() || clientName == rowClientReport["ClientName3"].ToString().ToUpper())
                                        {
                                            fileName = rowClientReport["Filename"].ToString();
                                            attachments.Rows.Add(fileName);
                                            logwrite.LogWrite("Below mentioned files are sent sucessfully sent" + Environment.NewLine + fileName + Environment.NewLine);
                                            fileName = "";
                                        }
                                        
                                       
                                    }

                                    if (attachments.Rows.Count>0)
                                    {
                                        MailMessage mail = new MailMessage();

                                        mail.From = new MailAddress(emailFrom);
                                        mail.Subject = subject;
                                        mail.Body = bodyFinal;
                                        mail.IsBodyHtml = true;
                                        foreach (DataRow rowAttachment in attachments.Rows)
                                        {

                                            mail.Attachments.Add(new Attachment(rowAttachment["AttachmentName"].ToString()));
                                        }
                                        mail.To.Add(emailTo);
                                        if(ccEmail!=string.Empty)
                                        { mail.CC.Add(ccEmail); }
                                        if (bccEmail != string.Empty)
                                        { mail.Bcc.Add(bccEmail); }
                                        using (SmtpClient smtp = new SmtpClient(smtpAddress, portNumber))
                                        {

                                            smtp.Credentials = new NetworkCredential(emailFrom, password);
                                            smtp.EnableSsl = enableSSL;
                                            smtp.Send(mail);

                                        }
                                        mail.Attachments.Clear();
                                        emailTo = "";
                                        attachments = null;
                                        ccEmail = "";
                                        bccEmail = "";
                                        
                                    }
                                    else 
                                    {
                                        
                                    }
                                    



                                    //   
                                  
                                                                                              
                             }
                               // runCount = runCount + 1;
                                logwrite.LogWrite("Mails sent susuccessfully!!!"+ Environment.NewLine + " files sucessfully sent" + Environment.NewLine + "Below mentioned files are sent sucessfully sent" + Environment.NewLine + fileName + Environment.NewLine + "Mail sent complete !!! ");

                                MessageBox.Show("Mails sent successfully");

                                Environment.Exit(-1);
                            
                        }
                        catch (Exception e)
                        {
                            logwrite.LogWrite(e.ToString());
                            // secound attempt to run the program
                            //if (runCount <= 2)
                            //{
                            //    sendMail();
                            //}
                            //else
                            //{
                            //    logwrite.LogWrite("Total run count="+ runCount);
                            //    runCount = 0;
                            //    Environment.Exit(-1);
                            //}
                                                        
                        }
                    }
                    else
                    {
                       
                        logwrite.LogWrite("Folder named "+filePath+" donot contain any files");
                        MessageBox.Show("Folder named " + filePath + " donot contain any files");
                        Environment.Exit(-1);
                    }
                }

                else
                {
                    logwrite.LogWrite("Folder named " +filePath+" not found");
                    MessageBox.Show("Folder named " + filePath + " not found");
                    Environment.Exit(-1);
                }
           // }


        }

    }
    public class LogWriter
    {
        private string m_exePath = string.Empty;
        public void LogWrite(string logMessage)
        {
            string path = @"D:\mailtest";
            m_exePath = path;
            try
            {
                using (StreamWriter w = File.AppendText(m_exePath + "\\" + "log.txt"))
                {
                    Log(logMessage, w);
                }
            }
            catch (Exception ex)
            {
            }
        }

        public void Log(string logMessage, TextWriter txtWriter)
        {
            try
            {
                txtWriter.Write("\r\nLog Entry : ");
                txtWriter.WriteLine("{0} {1}", DateTime.Now.ToLongTimeString(),
                    DateTime.Now.ToLongDateString());
               // txtWriter.WriteLine("  :");
                txtWriter.WriteLine(logMessage);
                txtWriter.WriteLine("-------------------------------");
            }
            catch (Exception ex)
            {
            }
        }
    }
    }


