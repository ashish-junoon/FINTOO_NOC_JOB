using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using LMS_DL;
using NLog;

namespace LOAN_CLOSER_APPLICATION
{
    class Program
    {
        private static readonly NLog.Logger logger = LogManager.GetCurrentClassLogger();
        static void Main()
        {
            string RootPath = ConfigurationManager.AppSettings["RootDirectory"].ToString();
            logger.Info($"Application started...{DateTime.Now}");
            try
            {
                 string rootPath = Directory.GetParent(AppDomain.CurrentDomain.BaseDirectory).Parent.Parent.FullName;
                string[] folders = {
               Path.Combine(RootPath, "LoanCloserDocument"),
            };
                DeleteFiles(folders);
                RunMainCode();
            }
            catch (Exception ex)
            {
                //logger.Error($"Error while Executing code - {ex.Message}");
                Console.WriteLine($"Unexpected error in Main method: {ex.Message}");
            }
            logger.Info($"Application end...{DateTime.Now}");
        }

        static void DeleteFiles(string[] folders)
        {
            try
            {
                foreach (var folder in folders)
                {
                    if (Directory.Exists(folder))
                    {
                        foreach (var file in Directory.GetFiles(folder, "*.html"))
                        {
                            File.Delete(file);
                        }
                        foreach (var file in Directory.GetFiles(folder, "*.pdf"))
                        {
                            File.Delete(file);
                        }
                    }
                    else
                    {
                        logger.Error($"Folder does not exist - {folder}");
                    }
                }
                logger.Error($"previoius .html and .pdf file deleted !!");
            }
            catch (Exception ex)
            {
                logger.Error($"Error while deleting files from folder: {ex.Message}");
            }
        }

        static void RunMainCode()
        {
            try
            {
                string connectionString = ConfigurationManager.ConnectionStrings["CrediCash_Dev"].ConnectionString;
                if (!string.IsNullOrEmpty(connectionString))
                {
                    List<InformationModel> informationModels = GetIncompleteData(connectionString);
                    logger.Info($"Data Count from Database: {informationModels.Count}");
                    if (informationModels.Count > 0)
                    {
                        foreach (var information in informationModels)
                        {
                            if (!string.IsNullOrEmpty(information.loan_id) && !string.IsNullOrEmpty(information.lead_id))
                            {
                                bool combinedServiceEnabled = Convert.ToBoolean(ConfigurationManager.AppSettings["LoanCloserService"]);
                                if (combinedServiceEnabled)
                                {
                                    if (!information.NOC_sent_over_email)
                                    {
                                        logger.Info($"Send loan closure process start: {information.loan_id}");
                                        LoanCloser.LoanCloserProcedure(connectionString, information);
                                        logger.Info($"Send loan closure process end: {information.loan_id}");
                                    }
                                    else
                                    {
                                        logger.Info($"NO any data found !!");
                                    }
                                }
                                else
                                {
                                    logger.Info($"Send indivisual process start");
                                }
                                System.Threading.Thread.Sleep(2000);
                            }
                        }
                    }
                    else
                    {
                        logger.Error($"Information: NO any data found for sending email.{informationModels.Count}");
                    }
                }
                else
                {
                    logger.Error($"Connection string is empty or missing..{connectionString}");
                }
            }
            catch (Exception ex)
            {
                logger.Error($"Error while running main code logic: {ex.Message}");
            }
        }
      
        public static List<InformationModel> GetIncompleteData(string dbconnection)
        {
            List<InformationModel> informationModels = new List<InformationModel>();
            DataSet objDs = null;

            try
            {
                using (var connection = new SqlConnection(dbconnection))
                {
                    SqlParameter[] param = new SqlParameter[1];
                    objDs = SqlHelper.ExecuteDataset(dbconnection, CommandType.StoredProcedure, "USP_GetIncompleteData_NOC", param);
                }

                if (objDs?.Tables[0] != null && objDs.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow row in objDs.Tables[0].Rows)
                    {
                        InformationModel information = new InformationModel
                        {
                            NOC_sent_over_email = row["noc_send_over_email"] != DBNull.Value && Convert.ToBoolean(row["noc_send_over_email"]),
                            loan_id = row["loan_id"] != DBNull.Value ? Convert.ToString(row["loan_id"]) : string.Empty,
                            lead_id = row["lead_id"] != DBNull.Value ? Convert.ToString(row["lead_id"]) : string.Empty,
                        };
                        informationModels.Add(information);
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error($"Error fetching incomplete data: {ex.Message}");
            }
            return informationModels;
        }
    
    }
}
