using iTextSharp.text;
using iTextSharp.text.html.simpleparser;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using Pechkin;
using Pechkin.Synchronized;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Text.RegularExpressions;


namespace PdfCutterService
{
    class PdfManager
    {
        public List<string> GetFileList(string type, string path)
        {
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            List<string> filePathList = Directory.GetFiles(path, type).ToList();
            return filePathList;
        }


        public void ExtractPdf(string type)
        {
            try
            {
                string inputPdfPath = ConfigurationManager.AppSettings["pdfFolderPatrh"];
                string processedPdfPath = ConfigurationManager.AppSettings["processedFile"];
                string outputPdfPath = ConfigurationManager.AppSettings["outputPdfPath"];
                if (!Directory.Exists(outputPdfPath))
                    Directory.CreateDirectory(outputPdfPath);

                List<string> inputPdfList = this.GetFileList(type, inputPdfPath);
                List<string> processedPdfList = this.GetFileList(type, processedPdfPath);



                foreach (string pdfFile in inputPdfList)
                {

                    string inputFileName = pdfFile.Replace(inputPdfPath, processedPdfPath)+DateTime.Now;
                    if (processedPdfList.Contains(inputFileName))
                        continue;

                    string textContent = this.ReadPdfFile(pdfFile);



                    List<string> contentList = this.GetContentList(textContent, "Message Header", "End of Message");

                    foreach (string content in contentList)
                    {
                        string outputFileName = this.GetSubstring(content, "Documentary Credit Number", @"F31C:").Replace("<br>", "").Trim();
                        string outputFileFullPath = @outputPdfPath + "\\" + outputFileName + "@0.pdf";
                        //ExportToPdf(FormatContent(headerText, content));
                        ExportToPdfByItext(this.FormatContent(content), outputFileFullPath);
                        this.WriteToFile(outputFileName + " \n file is splitted");
                    }
                    if (!Directory.Exists(processedPdfPath))
                        Directory.CreateDirectory(processedPdfPath);


                    File.Move(pdfFile, inputFileName);


                }
            }
            catch (Exception ex)
            {
                this.SendMail("MT700 Error Message:\n" + ex.Message, "MT700 Service Alert");
                throw ex;
            }

        }

        public string ReadPdfFile(string fileName)
        {
            StringBuilder text = new StringBuilder();

            if (File.Exists(fileName))
            {
                PdfReader pdfReader = new PdfReader(fileName);

                for (int page = 1; page <= pdfReader.NumberOfPages; page++)
                {
                    ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();
                    string currentText = PdfTextExtractor.GetTextFromPage(pdfReader, page, strategy);

                    currentText = Encoding.UTF8.GetString(ASCIIEncoding.Convert(Encoding.Default, Encoding.UTF8, Encoding.Default.GetBytes(currentText)));
                    text.Append(currentText);
                }
                pdfReader.Close();
            }
            return text.ToString();
        }

        public List<int> AllIndexesOf(string str, string value)
        {
            if (String.IsNullOrEmpty(value))
                throw new ArgumentException("the string to find may not be empty", "value");
            List<int> indexes = new List<int>();
            for (int index = 0; ; index += value.Length)
            {
                index = str.IndexOf(value, index);
                if (index == -1)
                    return indexes;
                indexes.Add(index);
            }
        }


        //format whole content
        public string FormatContent(string fullContent)
        {
            fullContent = fullContent.Replace("\n", "<br>");
            string[] pageNumber = Regex.Matches(fullContent, @"Page\s(\d|\d\d)\sof\s(\d\d|\d)").Cast<Match>().Select(m => m.Value).ToArray();
            for (int i = 0; i < pageNumber.Length; i++)
            {
                fullContent = fullContent.Replace(pageNumber[i], "");
            }
            StringBuilder sb = new StringBuilder();
            sb.Append("<table cellspacing='0' cellpadding='0'>");

            sb.Append("<br>");
            sb.Append("<tr bgcolor='#dbdbd6' style='font-size:8px; font-weight:bold' > <td>Message Header</td></tr>");
            sb.Append("</table>");

            sb.Append("<br>");
            string messageHeaderText = this.GetSubstring(fullContent, "Message Header", "Message Text");
            messageHeaderText = messageHeaderText.Replace(":", "");
            string[] messageHeaderTextkeywords = { "Swift Input", "Swift Output", "Sender", "Receiver", "MUR" };

            sb.Append(this.SetTabularContent(messageHeaderText, messageHeaderTextkeywords));
            sb.Append("<table cellspacing='0' cellpadding='0'>");
            sb.Append("<tr bgcolor='#dbdbd6' style='font-size:8px; font-weight:bold' > <td>Message Text</td></tr>");
            sb.Append("</table>");
            sb.Append("<br>");
            string messageText = "";
            if (fullContent.IndexOf("Message Trailer") != -1)
            {
                messageText = this.GetSubstring(fullContent, "Message Text", "Message Trailer");
            }
            else
            {
                messageText = fullContent.Substring(fullContent.IndexOf("Message Text") + "Message Text".Length);
            }
            string[] messageTextKeywords = Regex.Matches(messageText, @"F\S+:").Cast<Match>().Select(m => m.Value.Trim()).Where(m => m.Length <= 5 && m.Length > 3).ToArray();
            string[] numbering = Regex.Matches(messageText, @"\s\d\.").Cast<Match>().Select(m => m.Value).ToArray();
            for (int i = 0; i < numbering.Length; i++)
            {
                messageText = messageText.Replace(numbering[i], "<br>" + numbering[i]);
            }

            string[] numbering2 = Regex.Matches(messageText, @"\s\d\d\)").Cast<Match>().Select(m => m.Value).ToArray();
            for (int i = 0; i < numbering2.Length; i++)
            {
                messageText = messageText.Replace(numbering2[i], "<br>" + numbering2[i]);
            }

            string[] numbering3 = Regex.Matches(messageText, @"\s\S\)").Cast<Match>().Select(m => m.Value).ToArray();
            for (int i = 0; i < numbering3.Length; i++)
            {
                messageText = messageText.Replace(numbering3[i], "<br>" + numbering3[i]);
            }

            sb.Append(this.SetTabularContent(messageText, messageTextKeywords));

            sb.Append("<br>");

            if (fullContent.IndexOf("Message Trailer") != -1)
            {
                sb.Append("<table cellspacing='0' cellpadding='0'>");

                sb.Append("<br>");
                sb.Append("<tr bgcolor='#dbdbd6' style='font-size:8px; font-weight:bold;font-family:georgia,garamond,serif;' > <td>Message Trailer</td></tr>");
                sb.Append("</table>");
                string trailerMessageText = fullContent.Substring(fullContent.IndexOf("Message Trailer") + "Message Trailer".Length);
                sb.Append("<table id='mytable' width='100%' cellspacing='0' cellpadding='1'>");
                sb.Append("<tr style='font-size:9px; font-weight:normal; font-family:georgia,garamond,serif;'><td>");
                sb.Append(trailerMessageText);
                sb.Append("</td></tr>");
                sb.Append("</table>");
                sb.Append("<br>");
            }
            sb.Append("<table cellspacing='0' cellpadding='0'>");

            sb.Append("<br>");
            sb.Append("<tr bgcolor='#dbdbd6' style='font-size:8px; font-weight:bold; font-family:georgia,garamond,serif;' > <td>End Of Message</td></tr>");
            sb.Append("</table>");



            return sb.ToString();
        }

        public string GetSubstring(string content, string firstStr, string lastStr)
        {
            if (content.Contains(firstStr) && content.Contains(lastStr))
            {
                int firstIndex = content.IndexOf(firstStr) + firstStr.Length;
                int lastIndex = content.IndexOf(lastStr);
                return content.Substring(firstIndex, lastIndex - firstIndex);
            }
            else
            {
                return content;
            }
        }

        //create pdf by itextsharp

        public void ExportToPdfByItext(string content, string outputFilePath)
        {
            Document pdfDoc = new Document(PageSize.A4);

            //Create a New instance of PDFWriter Class for Output File

            PdfWriter writer = PdfWriter.GetInstance(pdfDoc, new FileStream(outputFilePath, FileMode.Create));
            PageEventHelper pageEventHelper = new PageEventHelper();
            writer.PageEvent = pageEventHelper;


            HTMLWorker htmlparser = new HTMLWorker(pdfDoc);

            //Open the Document

            pdfDoc.Open();



            htmlparser.Parse(new StringReader(content));
            //Add the content of Text File to PDF File

            // pdfDoc.Add(new Paragraph(content));

            //Close the Document

            pdfDoc.Close();




        }

        //for testing by pechkin
        public void ExportToPdf(string content, string outputpdfFullPath)
        {
            // Simple PDF from String


            byte[] pdfBuffer = new SynchronizedPechkin(new GlobalConfig()).Convert(new ObjectConfig()
                    .SetLoadImages(true).SetZoomFactor(1.7)
                    .SetPrintBackground(true)
                    .SetScreenMediaType(true)
                    .SetCreateExternalLinks(true), content);



            ByteArrayToFile(outputpdfFullPath, pdfBuffer);


        }

        //save pechkin pdf file //test
        public bool ByteArrayToFile(string _FileName, byte[] _ByteArray)
        {
            try
            {
                // Open file for reading
                FileStream _FileStream = new FileStream(_FileName, FileMode.Create, FileAccess.Write);
                // Writes a block of bytes to this stream using data from  a byte array.
                _FileStream.Write(_ByteArray, 0, _ByteArray.Length);

                // Close file stream
                _FileStream.Close();

                return true;
            }
            catch (Exception _Exception)
            {
                Console.WriteLine("Exception caught in process while trying to save : {0}", _Exception.ToString());
            }

            return false;
        }
        public List<string> GetContentList(string textContent, string firstString, string lastString)
        {
            List<int> firstIndex = this.AllIndexesOf(textContent, firstString);
            List<int> secoundIndex = this.AllIndexesOf(textContent, lastString);

            List<string> contentList = new List<string>();
            for (int i = 0; i < firstIndex.Count; i++)
            {
                string content = textContent.Substring(firstIndex[i], secoundIndex[i] - firstIndex[i]);
                contentList.Add(content);
            }
            return contentList;
        }

        //setting content in cell wise
        public string SetTabularContent(string content, string[] contentKeywords)
        {
            StringBuilder sb = new StringBuilder();
            string columnWidth = @"20%";
            string pre = "", font = "8px";
            if (contentKeywords.Length > 6)
            {
                pre = "<pre>";
                columnWidth = @"10%";
                font = "8px";
            }

            sb.Append("<table cellspacing=0 cellpadding=0  width=100% >" + pre);

            for (int i = 0; i < contentKeywords.Length; i++)
            {
                if (content.IndexOf(contentKeywords[i]) == -1)
                    continue;

                if (i == contentKeywords.Length - 1)
                {
                    sb.Append("<tr style='font-size:" + font + "; font-style:normal;'><td  width=" + columnWidth + "  valign='top'> " + contentKeywords[i].Replace(":", "") + @":</td><td>");
                    sb.Append(this.FormatMessageText(content.Substring(content.IndexOf(contentKeywords[i])).Replace(contentKeywords[i], "")));
                    //sb.Append("</td><td>");
                }
                else if (content.IndexOf(contentKeywords[i + 1]) == -1)
                {
                    int j = 1, t = 1;
                    while (content.IndexOf(contentKeywords[i + j]) == -1)
                    {
                        j++;
                        if ((i + j >= contentKeywords.Length - 1))
                        {
                            t = 0;
                            break;

                        }
                    }
                    if (t == 1)
                    {
                        sb.Append("<tr style='font-size:" + font + "; font-style:normal;'><td  width=" + columnWidth + " valign='top'> " + contentKeywords[i].Replace(":", "") + @":</td><td>");
                        sb.Append(this.FormatMessageText(this.GetSubstring(content, contentKeywords[i], contentKeywords[i + j])));
                    }
                    else
                    {
                        sb.Append("<tr style='font-size:" + font + "; font-style:normal;'><td  width=" + columnWidth + "  valign='top'> " + contentKeywords[i].Replace(":", "") + @":</td><td>");
                        sb.Append(this.FormatMessageText(content.Substring(content.IndexOf(contentKeywords[i])).Replace(contentKeywords[i], "")));
                    }

                }
                else
                {
                    sb.Append("<tr style='font-size:" + font + "; font-style:normal;'><td  width=" + columnWidth + " valign='top'> " + contentKeywords[i].Replace(":", "") + @":</td><td>");
                    sb.Append(this.FormatMessageText(this.GetSubstring(content, contentKeywords[i], contentKeywords[i + 1])));
                    // sb.Append("</td><td>");

                }
                sb.Append(pre + "</td></tr>");


            }
            sb.Append(pre + "</table>");
            return sb.ToString();
        }


        //insert newline after random keyword 
        public string FormatMessageText(string content)
        {
            string table = @"<table cellspacing='0' cellpadding='0'  width=100% align='left'><tr><td width=10%></td><td>";
            string[] keyWords ={
                                  "Sequence of Total","Form of Documentary Credit","Documentary Credit Number","Date of Issue",
                                  "Applicable Rules","Date and Place of Expiry","Applicant Bank - Party Identifier - Identifier Code",
                                  "Applicant","Beneficiary","Name and Address:","Currency Code, Amount","Available With ... By ... - Name and Address - Code",
                                  "Drafts at ...","Drawee - Party Identifier - Identifier Code","Identifier Code:","Partial Shipments","Transhipment",
                                  "Port of Loading/Airport of Departure","Port of Discharge/Airport of Destination","Latest Date of Shipment",
                                  "Description of Goods and/or Services","Documents Required","Additional Conditions","Charges","Period for Presentation in Days",
                                  "Confirmation Instructions","Instructions to the Paying/Accepting/Negotiating Bank","Sender to Receiver Information",
                                  "Place of Taking in Charge/Dispatch from .../Place of Receipt",@"'Advise Through' Bank - Party Identifier - Name and Address"
                              };
            for (int i = 0; i < keyWords.Length; i++)
            {
                if (content.Contains(keyWords[i]))
                {
                    if (content.Contains("Total:"))
                        content = content.Insert(content.IndexOf("Total:"), "<br> &nbsp;");

                    content = content.Insert(content.IndexOf(keyWords[i]) + keyWords[i].Length, table);
                    content = content.Insert(content.Length, "</td><tr></table>");


                }


            }
            return content;
        }

        public void WriteToFile(string text)
        {
            string folderPath = ConfigurationManager.AppSettings["ServiceLog"].ToString();
            if (!Directory.Exists(folderPath))
                Directory.CreateDirectory(folderPath);

            string path = folderPath + "\\MT700_PdfCutter_ServiceLog.txt";
            using (StreamWriter writer = new StreamWriter(path, true))
            {
                writer.WriteLine(string.Format(text, DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt")));
                writer.Close();
            }
        }



        //for sending mail
        public void SendMail(string body, string subject)
        {
            try
            {
                SmtpClient client = new SmtpClient();
                client.Port = 25;
                client.Host = "hocs01.southeastbank.com.bd";
                client.EnableSsl = true;
                client.Timeout = 10000;
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.UseDefaultCredentials = false;
                MailMessage mm = new MailMessage("abdulla.mamun@southeastbank.com.bd", "uzzal.koiri@southeastbank.com.bd", subject, body);
                mm.CC.Add("abdulla.mamun@southeastbank.com.bd");
                mm.BodyEncoding = UTF8Encoding.UTF8;
                mm.DeliveryNotificationOptions = DeliveryNotificationOptions.OnFailure;

                client.Send(mm);
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
    }
}
