using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Text.RegularExpressions;
namespace LOAN_CLOSER_APPLICATION
{
    public class AllEmailBody
    {
        private static readonly NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();
        public static string CombinedSingleEmailTemplate(string method_name, string name, string loan_id, string documentname)
        {
            string dynamicContent = "";
            string loanstatus = "";
            string customer_name = name;
            string loan_account_number = loan_id;
            string _documentname = documentname;
            if (_documentname == "_Loan_ForeClosure")
            {
                loanstatus = "ForeClosure";
                dynamicContent = "<ul>" +
                                 "   <li>Loan " + loanstatus + " – Loan " + loanstatus + " certificate .</li>" +
                                 "</ul>";
            }
            if (_documentname == "_NOC_Loan_Settlement")
            {
                loanstatus = "Settled";
                dynamicContent = "<ul>" +
                                 "   <li>Loan " + loanstatus + " – Loan " + loanstatus + " certificate .</li>" +
                                 "</ul>";
            }
            else
            {
                loanstatus = "closed";
                dynamicContent = "<ul>" +
                                 "   <li>Loan Closure – Loan Closure certificate .</li>" +
                                 "</ul>";
            }
            return (
                " <!DOCTYPE html>" +
                " <html>" +
                "   <body style='font-family: Arial, sans-serif; line-height: 1.5; color: #333; margin: 0; padding: 0;'>" +
                "       <div style='margin: 20px auto; padding: 20px;'>" +
                "           <div style='font-size: 18px; font-weight: bold; margin-bottom: 20px;'>Dear " + customer_name + ",</div>" +
                "               <p>Greetings from " + ConfigurationManager.AppSettings["company_product_name"] + "!</p>" +
                "               <p>We are pleased to inform you that your loan account <span style='font-weight: bold;'>" + loan_account_number + "</span> has been successfully " + loanstatus + ".</p>" +
                "               <p>Please find the attached documents for your reference:</p>" +
                                dynamicContent +
                "               <p>We sincerely thank you for placing your trust in <a href='" + ConfigurationManager.AppSettings["website_link"] + "'>" + ConfigurationManager.AppSettings["product_name"] + ".</a></p>" +
                "               <p> For any support or clarification you may require, please feel free to reach out to us at " + ConfigurationManager.AppSettings["MobileNo"] + ".</p>" +
                "               <p>We're always here to help you.<br>" +
                "               <p>Best Regards,<br>" +
                "               <strong>Team " + ConfigurationManager.AppSettings["company_product_name"] + "</strong><br>" +
                "               <strong>(Lending Partner : " + ConfigurationManager.AppSettings["company_name"] + ")</strong></p>" +
                "               <p style='font-size: 14px; color: #666; border-top: 1px solid #ddd; padding-top: 10px;'>" +
                "               <em>This is an auto-generated email. Please do not reply to this address.</em>" +
                "           </p>" +
                "       </div>" +
                "   </body>" +
                " </html>"
            );
        }


        public static List<string> SendEmail(List<string> pdfAttachmentPaths, string recipientEmail, string user_id, string lead_id , string loan_id , string disbursalProcedure, string emailBody, string subject, string name ,  string ProductCode = "PU")
        {
            List<string> successfullyAttachedFiles = new List<string>();
            try
            {
                string smtpServer = ConfigurationManager.AppSettings["SmtpServer"];
                int smtpPort = int.Parse(ConfigurationManager.AppSettings["SmtpPort"]);
                string senderEmail = ConfigurationManager.AppSettings["EmailFrom"];
                string senderPassword = ConfigurationManager.AppSettings["EmailPassword"];
                string EmailCC = ConfigurationManager.AppSettings["Emailcc"];
                string EmailBcc = ConfigurationManager.AppSettings["EmailBcc"];
                string htmlFilePath = Path.GetTempFileName() + ".html";
                MailMessage mailMessage = new MailMessage
                {
                    From = new MailAddress(senderEmail, ConfigurationManager.AppSettings["company_product_name"]),
                    Subject = subject,
                    Body = emailBody,
                    IsBodyHtml = true
                };

                if(Convert.ToBoolean(ConfigurationManager.AppSettings["IsDevelopment"]))
                {
                    mailMessage.To.Add(new MailAddress(ConfigurationManager.AppSettings["DevelopmentEmail"], name));
                }
                else
                {
                    mailMessage.To.Add(new MailAddress(recipientEmail, name));
                    if (!string.IsNullOrEmpty(EmailCC))
                    {
                        foreach (var email in EmailCC.Split('|'))
                        {
                            mailMessage.CC.Add(email);
                        }
                    }
                    if (!string.IsNullOrEmpty(EmailBcc))
                    {
                        foreach (var email in EmailBcc.Split('|'))
                        {
                            mailMessage.Bcc.Add(email);
                        }
                    }
                }
                foreach (var pdfPath in pdfAttachmentPaths)
                {
                    if (!string.IsNullOrEmpty(pdfPath) && File.Exists(pdfPath))
                    {
                        mailMessage.Attachments.Add(new Attachment(pdfPath));
                        successfullyAttachedFiles.Add(pdfPath);
                    }
                }
                
                SmtpClient smtpClient = new SmtpClient(smtpServer, smtpPort)
                {
                    Credentials = new NetworkCredential(senderEmail, senderPassword),
                    EnableSsl = true
                };

                mailMessage.DeliveryNotificationOptions = DeliveryNotificationOptions.OnSuccess | DeliveryNotificationOptions.OnFailure;

                try
                {
                    smtpClient.Send(mailMessage);
                    logger.Info("Email sent successfully!");
                    successfullyAttachedFiles.Add("success");

                }
                catch(SmtpFailedRecipientsException ex)
                {
                    logger.Info("Email not sent SmtpFailedRecipientsException!" + ex.Message);
                    successfullyAttachedFiles.Add("SmtpFailedRecipientsException");
                }
                catch(SmtpException ex)
                {
                    logger.Info("Email not sent SmtpException!" + ex.Message);
                    successfullyAttachedFiles.Add("SmtpException");
                }
                catch (Exception smtpEx)
                {
                    logger.Error($"SMTP error: {smtpEx.Message}");
                    successfullyAttachedFiles.Add("Failed !!");
                }
                
            }
            catch (Exception ex)
            {
                logger.Error($"Error setting up email: {ex.Message}");
            }

            return successfullyAttachedFiles;
        }
        public static void DispatchEmail(string loan_id , string loancloserpath, string emailId, string userId, string leadId, string name , string disbursalProcedure , string documentname)
        {
            List<string> pdfAttachmentPaths = new List<string>();
            if (!string.IsNullOrEmpty(loancloserpath)) pdfAttachmentPaths.Add(loancloserpath);
            string emailBody = CombinedSingleEmailTemplate(disbursalProcedure , name , loan_id , documentname);
            string subject = documentname == "_NO_DUES_CERTIFICATE" ? ConfigurationManager.AppSettings["LoanCloserSubject"].ToString()
                : documentname == "_Loan_ForeClosure" ? ConfigurationManager.AppSettings["LoanForeCloserSubject"].ToString()
                : documentname == "_NOC_Loan_Settlement" ? ConfigurationManager.AppSettings["LoanSettlementSubject"].ToString() : ConfigurationManager.AppSettings["LoanCloserSubject"].ToString();

            // ***************************** Send Email Single pdf and multiple pdf ****************************************

            List<string> sentAttachments = SendEmail(pdfAttachmentPaths, emailId, userId, leadId, loan_id, disbursalProcedure, emailBody, subject , name);
            foreach (var attachment in sentAttachments)
            {
                if (!string.IsNullOrEmpty(attachment))
                {
                    // ***************************** Find Folder Name and update the database ****************************************

                    Match match = Regex.Match(attachment, @"\\(LoanCloserDocument)\\", RegexOptions.IgnoreCase);
                    if (match.Success)
                    {
                        string folderName = match.Groups[1].Value;
                        if (folderName == "LoanCloserDocument")
                        {
                            LoanCloser.NOCupdate(userId, leadId , loan_id);
                            logger.Info("Data update Successfully !!");
                        }
                    }
                }
            }

        }
    }
}
