using System;
using System.Configuration;
using System.Collections.Specialized;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Net;
using System.Net.Http;
using System.Text.RegularExpressions;
using Outlook = Microsoft.Office.Interop.Outlook;


namespace NewForma_Log_Command
{
    class Program
    {
        static int Main(string[] args)
        {
            string ProjectPath = ConfigurationManager.AppSettings.Get("ProjectPath");
            if (!Directory.Exists(ProjectPath))
            {
                Console.WriteLine("S Drive not accessible");
                Console.ReadLine();
                return 1;
            }

            Console.WriteLine("Please select what Outlook folder to analyze. Type P for Procore or N for Newforma.");
            string EmailFolder = Console.ReadLine();

            if (EmailFolder.ToLower() == "p")
            {
                Console.WriteLine("Procore it is.");
                EmailFolder = "p";
                return PerryvilleProcoreFolderFunc();
            }
            else if (EmailFolder.ToLower() == "n")
            {
                Console.WriteLine("Newforma it is.");
                EmailFolder = "n";
                return PerryvilleNewformaFolderFunc();
            }
            else
            {
                Console.WriteLine("Input not recognized. Terminating");
                return 1;
            }

        }

        static void PostToSlack(string message) //used to be static async
        {
            var urlWithAccessToken = ConfigurationManager.AppSettings.Get("ProjectWebHook");
            Console.WriteLine(urlWithAccessToken.ToString());
            var client = new SlackClient(urlWithAccessToken);
            client.PostMessage(message);
        }


        static int PerryvilleNewformaFolderFunc()
        {
            string ProjectPath = ConfigurationManager.AppSettings.Get("ProjectPath");

            Outlook.Application myApp = new();

            Outlook.NameSpace mapiNameSpace = myApp.Session;

            Outlook.Folder myInbox = (Outlook.Folder)mapiNameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
            Outlook.Folder myPerryvilleNewformaFolder = (Outlook.Folder)myInbox.Folders["Perryville Newforma"];
            Outlook.Items oItems = myPerryvilleNewformaFolder.Items.Restrict("[UnRead] = true");

            Console.WriteLine("Items received: " + oItems.Count);

            //string xlPath = Path.Combine(ProjectPath, @"Perryville Project Spreadsheet.xlsx");

            string oFolderPath = "";

            IList<IList<object>> RFIData = new List<IList<object>>();
            IList<IList<object>> SubmittalData = new List<IList<object>>();
            List<Object> innerList = new();

            List<string> RFINumber = new();
            List<string> RFIDescription = new();
            List<string> RFIDate = new();

            List<string> SubmittalNumber = new();
            List<string> SubmittalDescription = new();
            List<string> SubmittalDate = new();

            //used to download the pdfs
            WebClient myWebClient = new();

            //check if there are unread emails in the folder
            if (oItems.Count == 0)
            {
                Console.WriteLine("No unread emails");
                return 0;
            }
            else
            {
                //Used to recap what happened at the end
                string RecapText = "";
                string RecapRFIs = "*RFIs:* \r\n";
                string RecapSubmittals = "*Submittals:* \r\n";

                List<Outlook.MailItem> oMsgList = new();

                for (int i = 1; i <= oItems.Count; i++)
                {
                    Outlook.MailItem oMsg = (Outlook.MailItem)oItems[i];
                    string oBody = oMsg.Body;
                    string oSubject = oMsg.Subject;
                    DateTime oDate = oMsg.ReceivedTime;
                    string oDateTime = oDate.ToString("d");
                    Console.WriteLine(oDateTime);

                    //filter unread emails for GWL ones
                    if (!oMsg.Subject.Contains("RE:") && !oMsg.Subject.Contains("FW:"))
                    {
                        if (oMsg.Subject.Contains("RFI Forwarded"))
                        {
                            //add the email to the email list to be marked as read later 
                            oMsgList.Add(oMsg);

                            //get the RFI number and subject of the RFI from the email body
                            string index0str = "Notification about RFI ";
                            int index0 = oBody.IndexOf(index0str) + index0str.Length;

                            string index1str = "A RFI has been forwarded.";
                            int index1 = oBody.IndexOf(index1str);

                            oSubject = oBody[index0..index1].Trim();

                            Console.WriteLine(oSubject);

                            int index = oBody.IndexOf("Sender ID: ");

                            string oRFINumber = "RFI " + oBody[(index + 11)..(index + 15)].Trim();
                            oFolderPath = oRFINumber + " (" + oSubject + ")";

                            Console.WriteLine(oFolderPath);

                            char[] separators = new char[] { '\\', '/', ':', '*', '?', '<', '>', '|' };

                            oFolderPath = oFolderPath.Replace(separators, "-");
                            oFolderPath = oFolderPath.Replace("\"", "''");

                            Console.WriteLine(oFolderPath);

                            //check if there is already a folder related to that RFI.If not, create it
                            if (!Directory.Exists(Path.Combine(ProjectPath, @"Project Information\Construction Admin\RFI", oFolderPath)))
                            {
                                Console.WriteLine($"RFI {i}: Creating Folder.");
                                Directory.CreateDirectory(Path.Combine(ProjectPath, @"Project Information\Construction Admin\RFI", oFolderPath));

                                //add the RFI to the Recap string
                                RecapRFIs += $"   - \"{oFolderPath}\" has been logged and downloaded. \r\n";

                                //add them to the lists for Excel
                                RFINumber.Add(oRFINumber);
                                RFIDescription.Add(oSubject);

                                //get the RFI pdf link from the email body
                                string oRFIDownloadFiles = Regex.Match(oBody[oBody.IndexOf("Download all files")..], @"\<([^>]*)\>").Groups[1].Value;
                                oRFIDownloadFiles = oRFIDownloadFiles.Replace("predownload", "download");

                                Console.WriteLine("Initiating file download.");
                                //download the RFI zip file and extract it
                                myWebClient.DownloadFile(oRFIDownloadFiles, Path.Combine(ProjectPath, @"Project Information\Construction Admin\RFI", oFolderPath, $"{oRFINumber}.zip"));
                                //myWebClient.DownloadFileAsync(new Uri(oRFIDownloadFiles, UriKind.Absolute), Path.Combine(ProjectPath, @"Project Information\Construction Admin\RFI", oFolderPath, $"{oRFINumber}.zip"));
                                ZipFile.ExtractToDirectory(Path.Combine(ProjectPath, @"Project Information\Construction Admin\RFI", oFolderPath, $"{oRFINumber}.zip"), Path.Combine(ProjectPath, @"Project Information\Construction Admin\RFI", oFolderPath));
                                Console.WriteLine("Download successful.");

                                //get the date of the RFI from the email body
                                string oRFIDate = oDateTime;  //Regex.Match(oBody[oBody.IndexOf("(Forwarded).pdf")..], @"[0-3]?[0-9]/[0-3]?[0-9]/(?:[0-9]{2})?[0-9]{2}").Groups[0].Value;

                                //add the date to the list of dates for Excel
                                RFIDate.Add(oRFIDate);

                                innerList.Clear();

                                innerList.Add(oRFINumber);
                                innerList.Add(oSubject);
                                innerList.Add(oRFIDate);

                                RFIData.Add(new List<object>(innerList));
                            }
                            else
                            {
                                //It is assumed that if the folder is there, then the files have been downloaded and logged already
                                Console.WriteLine($"RFI {i}: Folder already existing.");
                            } // end of existing folder
                        } //enf of rfi forwarded

                        else if (oMsg.Subject.Contains("Submittal Forwarded"))
                        {
                            //add the email to the email list to be marked as read later 
                            oMsgList.Add(oMsg);

                            //get the submittal number and subject of the submittal from the email subject

                            string index0str = "Notification about Submittal ";
                            int index0 = oBody.IndexOf(index0str) + index0str.Length;

                            string index1str = "A Submittal has been forwarded.";
                            int index1 = oBody.IndexOf(index1str);

                            oSubject = oBody[index0..index1].Trim();

                            Console.WriteLine(oSubject);

                            int index = oBody.IndexOf("Sender ID: ");

                            string oSubmittalNumber = oBody[(index + 11)..(index + 25)].Trim().Replace(" ", string.Empty);
                            oFolderPath = oSubmittalNumber + " " + oSubject;

                            oFolderPath = string.Join("-", oFolderPath.Split(Path.GetInvalidFileNameChars()));

                            //check if there is already a folder related to that submittal.If not, create it
                            if (!Directory.Exists(Path.Combine(ProjectPath, @"Project Information\Construction Admin\Submittals", oFolderPath)))
                            {
                                Console.WriteLine($"Submittal {i}: Creating Folder.");
                                Directory.CreateDirectory(Path.Combine(ProjectPath, @"Project Information\Construction Admin\Submittals", oFolderPath));

                                //add the submittal to the Recap string
                                RecapSubmittals += $"   - \"{oFolderPath}\" has been logged and downloaded. \r\n";

                                //add them to the lists for Excel
                                SubmittalNumber.Add(oSubmittalNumber);
                                SubmittalDescription.Add(oSubject);

                                //get the submittal pdf link from the email body
                                string oSubmittalDownloadFiles = Regex.Match(oBody[oBody.IndexOf("Download all files")..], @"\<([^>]*)\>").Groups[1].Value;
                                oSubmittalDownloadFiles = oSubmittalDownloadFiles.Replace("predownload", "download");

                                //download the submittal zip file and extract it
                                Console.WriteLine("Initiating file download.");
                                myWebClient.DownloadFile(oSubmittalDownloadFiles, Path.Combine(ProjectPath, @"Project Information\Construction Admin\Submittals", oFolderPath, $"{oSubmittalNumber}.zip"));
                                ZipFile.ExtractToDirectory(Path.Combine(ProjectPath, @"Project Information\Construction Admin\Submittals", oFolderPath, $"{oSubmittalNumber}.zip"), Path.Combine(ProjectPath, @"Project Information\Construction Admin\Submittals", oFolderPath));
                                Console.WriteLine("Download successful.");

                                //get the date of the Submittal from the email body
                                string oSubmittalDate = oDateTime; 

                                //add the date to the list of dates for Excel
                                SubmittalDate.Add(oSubmittalDate);

                                innerList.Clear();

                                innerList.Add(oSubmittalNumber);
                                innerList.Add("");
                                innerList.Add(oSubject);
                                innerList.Add(oSubmittalDate);

                                SubmittalData.Add(new List<object>(innerList));
                            }
                            else
                            {
                                //It is assumed that if the folder is there, then the files have been downloaded and logged already
                                Console.WriteLine($"Submittal {i}: Folder already existing.");
                            } // end of existing folder
                        } //enf of submittal forwarded
                    } //end of "email subject contains perryville"
                } //end of the unreal emails loop

                if (RFINumber.Count + SubmittalNumber.Count > 0)
                {
                    //ExcelClass.OpenExcel(xlPath, RFINumber, RFIDescription, RFIDate, SubmittalNumber, SubmittalDescription, SubmittalDate);
                    RecapText += "We received some new CA documents. I downloaded and logged them for you. \r\nHere is a recap of the changes: \r\n \r\n";
                }

                if (RFINumber.Count > 0)
                {
                    GoogleSheetsClass.WritetoGoogleSheets("RFIs", RFIData);
                    RecapText += RecapRFIs + "\r\n";
                }

                if (SubmittalNumber.Count > 0)
                {
                    GoogleSheetsClass.WritetoGoogleSheets("Submittals", SubmittalData);
                    RecapText += RecapSubmittals + "\r\n";
                }

                if (RFINumber.Count + SubmittalNumber.Count > 0)
                {
                    PostToSlack(RecapText);
                }

                Console.WriteLine(RecapText);

                //mark all the emails in the email list as read
                foreach (Outlook.MailItem oMsg in oMsgList)
                {
                    oMsg.UnRead = false;
                }

                return 0;
            }

        }


        static int PerryvilleProcoreFolderFunc()
        {
            string ProjectPath = ConfigurationManager.AppSettings.Get("ProjectPath");

            Outlook.Application myApp = new();

            Outlook.NameSpace mapiNameSpace = myApp.Session;

            var ProjectInboxFolder = ConfigurationManager.AppSettings.Get("ProjectInboxFolder");

            Outlook.Folder myInbox = (Outlook.Folder)mapiNameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
            Outlook.Folder myPerryvilleProcoreFolder = (Outlook.Folder)myInbox.Folders[ProjectInboxFolder];
            Outlook.Items oItems = myPerryvilleProcoreFolder.Items.Restrict("[UnRead] = true");

            Console.WriteLine("Items received: " + oItems.Count);

            string oFolderPath = "";

            //used for the excel/Google Sheets files

            IList<IList<object>> RFIData = new List<IList<object>>();
            IList<IList<object>> SubmittalData = new List<IList<object>>();
            List<Object> innerList = new();

            List<string> RFINumber = new();
            List<string> RFIDescription = new();
            List<string> RFIDate = new();

            List<string> SubmittalNumber = new();
            List<string> SubmittalDescription = new();
            List<string> SubmittalDate = new();

            //used to download the pdfs
            WebClient myWebClient = new();

            //check if there are unread emails in the folder
            if (oItems.Count == 0)
            {
                Console.WriteLine("No unread emails");
                return 0;
            }
            else
            {
                //Used to recap what happened at the end
                string RecapText = "";
                string RecapRFIs = "*RFIs:* \r\n";
                string RecapSubmittals = "*Submittals:* \r\n";

                List<Outlook.MailItem> oMsgList = new();

                for (int i = 1; i <= oItems.Count; i++)
                {
                    Outlook.MailItem oMsg = (Outlook.MailItem)oItems[i];
                    string oBody = oMsg.Body;
                    string oSubject = oMsg.Subject;
                    DateTime oDate = oMsg.ReceivedTime;
                    string oDateTime = oDate.ToString("d");
                    Console.WriteLine(oDateTime);

                    //filter unread emails for GWL ones
                    if (oMsg.Subject.Contains("Great Wolf Lodge - Perryville") )
                    {
                        if (oMsg.Subject.Contains("New RFI"))
                        {
                            //add the email to the email list to be marked as read later 
                            oMsgList.Add(oMsg);

                            oSubject = oSubject[oSubject.IndexOf("RFI")..];

                            int index = oSubject.IndexOf("(");
                            Console.WriteLine("( index = " + index.ToString());
                            string oRFINumber = oSubject[..(index-1)];

                            oFolderPath = oSubject;
                            Console.WriteLine(oSubject);
                            oSubject = oSubject[(index+1)..].TrimEnd(')');

                            Console.WriteLine(oRFINumber);
                            Console.WriteLine(oSubject);
                            Console.WriteLine(oFolderPath);

                            char[] separators = new char[] { '\\', '/', ':', '*', '?', '<', '>', '|' };

                            oFolderPath = oFolderPath.Replace(separators, "-");
                            oFolderPath = oFolderPath.Replace("\"", "''");

                            Console.WriteLine(oFolderPath);

                            //check if there is already a folder related to that RFI.If not, create it
                            if (!Directory.Exists(Path.Combine(ProjectPath, @"Project Information\Construction Admin\RFI", oFolderPath)))
                            {
                                Console.WriteLine($"RFI {i}: Creating Folder.");
                                Directory.CreateDirectory(Path.Combine(ProjectPath, @"Project Information\Construction Admin\RFI", oFolderPath));

                                //add the RFI to the Recap string
                                RecapRFIs += $"   - \"{oFolderPath}\" has been logged and downloaded. \r\n";

                                //add them to the lists for Excel
                                RFINumber.Add(oRFINumber);
                                RFIDescription.Add(oSubject);

                                //get the RFI pdf link from the email body
                                oBody = oBody[oBody.IndexOf("View PDF")..];
                                string oRFIViewPDF = Regex.Match(oBody, @"\<([^>]*)\>").Groups[1].Value;

                                //get the section relating to the attachments from the email body 
                                string oAttachments = oBody[oBody.IndexOf("Attachments:")..][14..];
                                int oAttEnd = oAttachments.IndexOf("Additional");
                                oAttachments = oAttachments.Substring(0, oAttEnd - 6);

                                //get the link to the attachment from the attachment section
                                char[] AttachmentdelimiterChars = { '<', '>' };
                                string oAttachmentName = "";
                                string oAttachmentLink = "";

                                //if the attachment section contains "<" then it means there is at least one attachment
                                if (oAttachments.Contains('<'))
                                {
                                    //split attachment section into attachment name/attachment link
                                    string[] oAttachmentSplits = oAttachments.Split(AttachmentdelimiterChars);

                                    foreach (string oSubString in oAttachmentSplits)
                                    {
                                        //if the substring index is pair, it's an attachment name
                                        if (Array.IndexOf(oAttachmentSplits, oSubString) % 2 == 0)
                                        {
                                            //remove trailing white spaces
                                            oAttachmentName = oSubString.Trim();
                                        }
                                        //if it is odd, then it's a link 
                                        else
                                        {
                                            //remove all white spaces
                                            oAttachmentLink = Regex.Replace(oSubString, @"\s+", "");

                                            //with the name and link, we can download and file the attachment
                                            myWebClient.DownloadFile(oAttachmentLink, Path.Combine(ProjectPath, @"Project Information\Construction Admin\RFI", oFolderPath, oAttachmentName));

                                            //add it to the RFI Recap
                                            RecapRFIs += $"              • Attachment {oAttachmentName} has been downloaded. \r\n";
                                        }
                                    }
                                }
                                else { Console.WriteLine($"RFI {i}: No Attachment."); }

                                myWebClient.DownloadFile(oRFIViewPDF, Path.Combine(ProjectPath, @"Project Information\Construction Admin\RFI", oFolderPath, oFolderPath) + ".pdf");

                                //get the date of the RFI from the email body
                                string oRFIDate = oDateTime; 

                                //add the date to the list of dates for Excel
                                RFIDate.Add(oRFIDate);

                                innerList.Clear();

                                innerList.Add(oRFINumber);
                                innerList.Add(oSubject);
                                innerList.Add(oRFIDate);

                                RFIData.Add(new List<object>(innerList));
                            }
                            else
                            {
                                //It is assumed that if the folder is there, then the files have been downloaded and logged already
                                Console.WriteLine($"RFI {i}: Folder already existing.");
                            } // end of existing folder
                        } //enf of rfi forwarded

                        else if (oMsg.Subject.Contains("Action Required for Submittal"))
                        {
                            //add the email to the email list to be marked as read later 
                            oMsgList.Add(oMsg);

                            //get the submittal number and subject of the submittal from the email subject
                            //get the submittal identity from the email subject
                            char[] SubjectdelimiterChars = { '(', ')' };
                            oSubject = oSubject.Split(SubjectdelimiterChars)[1].Trim()[1..];
                            string[] oSubjectSplit = oSubject.Split(':');
                            oSubject = Regex.Replace(oSubjectSplit[0], @"\s+", "") + oSubjectSplit[1];

                            char[] separators = new char[] { '\\', '/', ':', '*', '?', '<', '>', '|' };
                            oSubject = oSubject.Replace(separators, "-");
                            oSubject = oSubject.Replace("\"", "''");

                            string FolderPath = Path.Combine(ProjectPath, @"Project Information\Construction Admin\Submittals", oSubject);

                            //check if there is already a folder related to that submittal. If not, create it
                            if (!Directory.Exists(FolderPath))
                            {
                                Console.WriteLine($"Submittal {i}: Creating Folder.");
                                Directory.CreateDirectory(FolderPath);

                                //add the submittal to the Recap string
                                RecapSubmittals += $"   - \"{oSubject}\" has been logged but *NOT* downloaded. \r\n";

                                string oSubmittalNumber = Regex.Replace(oSubjectSplit[0], @"\s+", "");

                                //add them to the lists for Excel
                                SubmittalNumber.Add(oSubmittalNumber);
                                SubmittalDescription.Add(oSubjectSplit[1].Trim());

                                //get the RFI Procore link from the email body
                                oBody = oBody[oBody.IndexOf("View online")..];
                                string SubmittalUrl = Regex.Match(oBody, @"\<([^>]*)\>").Groups[1].Value;

                                CreateProcoreShortcut(FolderPath, "Procore Shortcut", SubmittalUrl);
                                Console.WriteLine("Shortcut Created.");

                                //get the date of the Submittal from the email body
                                string oSubmittalDate = oDateTime; 

                                //add the date to the list of dates for Excel
                                SubmittalDate.Add(oSubmittalDate);

                                innerList.Clear();

                                innerList.Add(oSubmittalNumber);
                                innerList.Add("");
                                innerList.Add(oSubjectSplit[1].Trim());
                                innerList.Add(oSubmittalDate);

                                SubmittalData.Add(new List<object>(innerList));
                            }
                            else
                            {
                                //It is assumed that if the folder is there, then the files have been downloaded and logged already
                                Console.WriteLine($"Submittal {i}: Folder already existing.");
                            } // end of existing folder
                        } //enf of submittal forwarded
                    } //end of "email subject contains perryville"
                } //end of the unreal emails loop

                if (RFINumber.Count + SubmittalNumber.Count > 0)
                {
                    RecapText += "We received some new CA documents. I logged them for you. \r\n *Please note submittals are NOT downloaded. Shortcuts to the Procore website were created instead.* \r\nHere is a recap of the changes: \r\n \r\n";
                }

                if (RFINumber.Count > 0)
                {
                    GoogleSheetsClass.WritetoGoogleSheets("RFIs", RFIData);
                    RecapText += RecapRFIs + "\r\n";
                }

                if (SubmittalNumber.Count > 0)
                {
                    GoogleSheetsClass.WritetoGoogleSheets("Submittals", SubmittalData);
                    RecapText += RecapSubmittals + "\r\n";
                }

                if (RFINumber.Count + SubmittalNumber.Count > 0)
                {
                    PostToSlack(RecapText);
                }

                Console.WriteLine(RecapText);

                //mark all the emails in the email list as read
                foreach (Outlook.MailItem oMsg in oMsgList)
                {
                    oMsg.UnRead = false;
                }

                return 0;
            }

        }

        static void CreateProcoreShortcut(string FolderPath, string FileName, string Url)
        {
            using (StreamWriter writer = new StreamWriter(FolderPath + @"\" + FileName + ".url"))
            {
                writer.WriteLine("[InternetShortcut]");
                writer.WriteLine("URL=" + Url);
                writer.Flush();
            }

        }

    }


    public static class ExtensionMethods
    {
        public static string Replace(this string s, char[] separators, string newVal)
        {
            string[] temp;

            temp = s.Split(separators, StringSplitOptions.RemoveEmptyEntries);
            return String.Join(newVal, temp);
        }
    }
}
