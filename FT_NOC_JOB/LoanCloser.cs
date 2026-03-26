using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Threading;
using iText.Html2pdf;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Layout;
using LMS_DL;
using Path = System.IO.Path;

namespace LOAN_CLOSER_APPLICATION
{
    public class LoanCloser
    {

        private static readonly NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();
        #region Loan Closer logic ----------------*********
        public static void LoanCloserProcedure(string connectionString, InformationModel information)
        {
            try
            {
                if (!string.IsNullOrEmpty(connectionString))
                {
                    LoanCloserModel loanCloserModels = ProcessNOCLetter(information, connectionString, ConfigurationManager.AppSettings["NOCService"]);
                    if (loanCloserModels != null)
                    {
                        string _rootPath = ConfigurationManager.AppSettings["RootDirectory"].ToString();
                        string _loancloserDocumentPath = Path.Combine(_rootPath, "LoanCloserDocument");
                        logger.Info($"disbursalDocumentPath path - {_loancloserDocumentPath}");
                        string _loancloserFilePath = Path.Combine(_loancloserDocumentPath, $"{loanCloserModels.product_name}{loanCloserModels.doucumentname}.txt");
                        if (File.Exists(_loancloserFilePath))
                        {
                            string loanclosertextContent = File.ReadAllText(_loancloserFilePath);
                            Dictionary<string, string> replacements = LoanCloserReplacementData(loanCloserModels);
                            foreach (var entry in replacements)
                            {
                                loanclosertextContent = loanclosertextContent.Replace(entry.Key, entry.Value);
                            }
                            string htmlContent = $@"{loanclosertextContent}";
                            string userFileName = loanCloserModels.loan_id.Replace(" ", "_");
                            string product_name = loanCloserModels.product_name.Replace(" ", "_");
                            string htmlFilePath = Path.Combine(_loancloserDocumentPath, $"{userFileName}_{product_name}{loanCloserModels.doucumentname}.html");
                            try
                            {
                                File.WriteAllText(htmlFilePath, htmlContent);
                            }
                            catch (Exception ex)
                            {
                                logger.Error($"Error: While writing file in directory!-  {ex.Message}");
                            }

                            try
                            {
                                NOC_SAVE(loanCloserModels.loan_id , loanCloserModels.user_id, loanCloserModels.lead_id, htmlContent);
                            }
                            catch (Exception ex)
                            {
                                logger.Error($"Error: While update data in DB for Loan id {information.loan_id} Error !-  {ex.Message}");
                            }

                            string loancloserpath = Path.Combine(_loancloserDocumentPath, $"{userFileName}_{product_name}{loanCloserModels.doucumentname}.pdf");
                            using (PdfWriter writer = new PdfWriter(loancloserpath))
                            using (PdfDocument pdfDoc = new PdfDocument(writer))
                            {
                                pdfDoc.SetDefaultPageSize(PageSize.A4);
                                Document document = new Document(pdfDoc);
                                document.SetMargins(40, 40, 40, 40);
                                ConverterProperties props = new ConverterProperties();
                                HtmlConverter.ConvertToPdf(new FileStream(htmlFilePath, FileMode.Open), pdfDoc, props);
                                document.Close();
                            }

                            AllEmailBody.DispatchEmail(loanCloserModels.loan_id, loancloserpath, loanCloserModels.email_id, loanCloserModels.user_id, loanCloserModels.lead_id, loanCloserModels.name, ConfigurationManager.AppSettings["LoanCloserMethodName"], loanCloserModels.doucumentname);

                            //Task.Run(() =>
                            //{
                            //    AllEmailBody.DispatchEmail(loanCloserModels.loan_id, loancloserpath, loanCloserModels.email_id, loanCloserModels.user_id, loanCloserModels.lead_id, loanCloserModels.name, ConfigurationManager.AppSettings["LoanCloserMethodName"], loanCloserModels.doucumentname);
                            //});

                        }
                        else
                        {
                            logger.Error($"Error: TXT file not found!-  {_loancloserFilePath}");
                        }
                    }

                    Thread.Sleep(5000);
                }
                else
                {
                    logger.Error("Please check connection string!");
                }
            }
            catch (Exception ex)
            {
                logger.Error("An error occurred: " + ex.Message);
            }
        }
        static Dictionary<string, string> LoanCloserReplacementData(LoanCloserModel loanCloser)
        {
            try
            {
                return new Dictionary<string, string>
                {
                    { "[name]", string.IsNullOrEmpty(loanCloser.name) ? "N/A" : loanCloser.name },
                    { "[loan_id]", string.IsNullOrEmpty(loanCloser.loan_id) ? "N/A" : loanCloser.loan_id },
                    { "[tenure]", string.IsNullOrEmpty(loanCloser.tenure) ? "N/A" : loanCloser.tenure },
                    { "[_tenure]", string.IsNullOrEmpty(loanCloser._tenure) ? "N/A" : loanCloser._tenure },
                    { "[email_id]", string.IsNullOrEmpty(loanCloser.email_id) ? "N/A" : loanCloser.email_id },
                    { "[loan_amount]", string.IsNullOrEmpty(loanCloser.loan_amount) ? "N/A" : loanCloser.loan_amount },
                    { "[address]", string.IsNullOrEmpty(loanCloser.address) ? "N/A" : loanCloser.address },
                    { "[current_date]", string.IsNullOrEmpty(loanCloser.current_date) ? "N/A" : loanCloser.current_date },
                    { "[closed_date]", string.IsNullOrEmpty(loanCloser.created_date) ? "N/A" : loanCloser.created_date },
                    { "[company_product_name]", ConfigurationManager.AppSettings["company_product_name"].ToString() },
                    { "[company_name]", ConfigurationManager.AppSettings["company_name"].ToString() },
                    { "[company_short_name]", ConfigurationManager.AppSettings["company_short_name"].ToString() },
                };
            }
            catch (Exception ex)
            {
                logger.Error($"Error in Loan Closer Replacement Data: {ex.Message}");
                return new Dictionary<string, string>();
            }
        }
        public static LoanCloserModel ProcessNOCLetter(InformationModel informationModels, string dbconnection, string method_name)
        {
            LoanCloserModel loanCloser = null;
            DataSet objDs = null;
            SqlParameter[] param = new SqlParameter[1];
            try
            {
                param[0] = new SqlParameter("loan_id", SqlDbType.VarChar, 30);
                param[0].Value = informationModels.loan_id;
                using (var connection = new SqlConnection(dbconnection))
                {
                    logger.Info("Fatch data in database !!");
                    objDs = SqlHelper.ExecuteDataset(dbconnection, CommandType.StoredProcedure, "USP_NOC_letter_EXE", param);
                    logger.Info("Fatch data in database !!");
                }
                if (objDs?.Tables[0] != null && objDs.Tables[0].Rows.Count > 0)
                {
                    DataRow row = objDs.Tables[0].Rows[0];

                    string _tenure = "0.00";
                    decimal total_tenure;
                    //MidpointRounding mode = MidpointRounding.ToEven;
                    string tenure = row["tenure"]?.ToString();
                    if (row["repayment_frequency"]?.ToString() == "0")
                    {
                        tenure = tenure.Replace("DAYS", "");
                        _tenure = tenure.Replace(tenure, "DAYS").Trim();
                    }
                    else
                    {
                        tenure = tenure.Replace("Months", "");
                        _tenure = tenure.Replace(tenure, "Months").Trim();
                    }

                    total_tenure = Convert.ToInt32(tenure);

                    loanCloser = new LoanCloserModel
                    {
                        user_id = row["user_id"]?.ToString() ?? "N/A",
                        lead_id = row["lead_id"]?.ToString() ?? "N/A",
                        name = row["full_name"]?.ToString() ?? "N/A",
                        current_date = row["current_date"] != DBNull.Value ? Convert.ToDateTime(row["current_date"]).ToString("yyyy-MM-dd") : DateTime.Now.ToString("yyyy-MM-dd"),
                        created_date = row["updated_date"] != DBNull.Value ? Convert.ToDateTime(row["updated_date"]).ToString("yyyy-MM-dd") : DateTime.Now.ToString("yyyy-MM-dd"),
                        tenure = tenure,
                        mobile_number = row["mobile_number"]?.ToString() ?? "NA",
                        email_id = row["email_id"]?.ToString() ?? "NA",
                        _tenure = _tenure.ToString() ?? "NA",
                        loan_id = row["loan_id"]?.ToString() ?? "NA",
                        product_name = row["product_name"]?.ToString() ?? "NA",
                        loan_amount = row["Loan_amount"]?.ToString() ?? "0",
                        address = row["address"]?.ToString() ?? "NA",
                        is_active = Convert.ToInt32(row["is_active"]),
                        doucumentname = Convert.ToInt32(row["is_active"]) == 10 ? "_NO_DUES_CERTIFICATE" : Convert.ToInt32(row["is_active"]) == 11 ? "_Loan_ForeClosure" :Convert.ToInt32(row["is_active"]) == 12 ? "_NOC_Loan_Settlement" : "UNKNOWN_DOCUMENT",
                    };
                    return loanCloser;

                }
            }
            catch (Exception ex)
            {
                string documentname = "Loan Closer";
                //InformationSendEmail(informationModels.loan_id, informationModels.full_name, informationModels.lead_id, documentname, ex.Message);
            }
            return loanCloser;


        }

        public static void NOCupdate(string user_id, string lead_id , string loan_id)
        {
            string connectionString = ConfigurationManager.ConnectionStrings["CrediCash_Dev"].ConnectionString;
            DataSet Objds = null;
            DataTable Objtable = new DataTable();
            SqlParameter[] param = new SqlParameter[4];

            param[0] = new SqlParameter("user_id", SqlDbType.VarChar, 10);
            param[0].Value = user_id;

            param[1] = new SqlParameter("lead_id", SqlDbType.VarChar, 10);
            param[1].Value = lead_id;

            param[2] = new SqlParameter("loan_id", SqlDbType.VarChar, 30);
            param[2].Value = loan_id;

            using (var connection = new SqlConnection(connectionString))
            {
                Objds = new DataSet();
                try
                {
                    Objds = SqlHelper.ExecuteDataset(connectionString, CommandType.StoredProcedure, "USP_update_NOC_letter_exe", param);
                }
                catch (Exception ex)
                {
                    logger.Error($"Error updateing Loan Closer Letters: {ex.Message}");
                }
            }
        }
        public static void NOC_SAVE(string loan_id , string user_id, string lead_id, string nochtml)
        {
            string connectionString = ConfigurationManager.ConnectionStrings["CrediCash_Dev"].ConnectionString;
            DataSet Objds = null;
            DataTable Objtable = new DataTable();
            SqlParameter[] param = new SqlParameter[4];

            param[0] = new SqlParameter("user_id", SqlDbType.VarChar, 10);
            param[0].Value = user_id;

            param[1] = new SqlParameter("lead_id", SqlDbType.VarChar, 10);
            param[1].Value = lead_id;

            param[2] = new SqlParameter("NOC_letter", SqlDbType.Text);
            param[2].Value = nochtml;

            param[3] = new SqlParameter("loan_id", SqlDbType.VarChar, 30);
            param[3].Value = loan_id;

            using (var connection = new SqlConnection(connectionString))
            {
                Objds = new DataSet();
                try
                {
                    Objds = SqlHelper.ExecuteDataset(connectionString, CommandType.StoredProcedure, "USP_save_NOC_letter_exe", param);
                }
                catch (Exception ex)
                {
                    logger.Error($"Error save Loan Closer Letters: {ex.Message}");
                }
            }
        }
        #endregion
    }
}
