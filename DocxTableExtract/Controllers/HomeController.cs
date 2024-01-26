using System;
using System.Collections.Generic;
using System.IO;
using System.Web;
using System.Web.Mvc;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace YourMvcApplication.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        private int CountTablesInDocument(WordprocessingDocument document)
        {
            int tableCount = 0;

            foreach (var table in document.MainDocumentPart.Document.Body.Elements<Table>())
            {
                tableCount++;
            }

            return tableCount;
        }

        private List<(int Index, string TableData)> GetTableDataList(WordprocessingDocument document)
        {
            var tableDataList = new List<(int Index, string TableData)>();

            int tableIndex = 1;

            foreach (var table in document.MainDocumentPart.Document.Body.Elements<Table>())
            {
                var tableData = "<table style='border-collapse: collapse; width: 100%; border: 4px solid black;'>";

                foreach (var row in table.Elements<TableRow>())
                {
                    tableData += "<tr style='border: 4px solid black;'>";

                    foreach (var cell in row.Elements<TableCell>())
                    {
                        tableData += $"<td style='border: 4px solid black; padding: 8px;'>{cell.InnerText}</td>";
                    }

                    tableData += "</tr>";
                }

                tableData += "</table>";
                tableDataList.Add((tableIndex, tableData));

                tableIndex++;
            }

            return tableDataList;
        }

        //[HttpPost]
        //public ActionResult AskQuestion(string question)
        //{
        //    if (!string.IsNullOrEmpty(question))
        //    {
        //        try
        //        {
        //            using (Stream stream = System.IO.File.Open(Server.MapPath("~/TempFiles/Cape - RFP Formal- PW 011015 (1) (1).docx"), FileMode.Open))
        //            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(stream, false))
        //            {
        //                // Placeholder logic: Retrieve information based on the question
        //                string answer = GetAnswerFromQuestion(wordDocument, question);
        //                return Content(answer);
        //            }
        //        }
        //        catch (Exception ex)
        //        {
        //            return Content("Error processing the document: " + ex.Message);
        //        }
        //    }
        //    else
        //    {
        //        return Content("Please provide a valid question.");
        //    }
        //}

        //private string GetAnswerFromQuestion(WordprocessingDocument document, string question)
        //{
        //    // Placeholder logic: Replace this with your actual implementation
        //    // For example, you can search for keywords in the document and provide relevant information

        //    // Convert the question to lowercase for case-insensitive matching
        //    string lowercaseQuestion = question.ToLower();

        //    // Placeholder keywords for demonstration purposes
        //    List<string> keywords = new List<string> { "asset management", "erp", "purchasing", "payroll", "hr" };

        //    // Check if any keyword matches the question
        //    foreach (var keyword in keywords)
        //    {
        //        if (lowercaseQuestion.Contains(keyword))
        //        {
        //            // Return a simple response for demonstration
        //            return $"The document contains information about {keyword}.";
        //        }
        //    }

        //    // Return a generic response if no specific keyword is found
        //    return "Sorry, I couldn't find relevant information based on your question.";
        //}


        private void SaveTablesToFolder(List<(int Index, string TableData)> tableDataList, string folderPath)
        {
            foreach (var tableData in tableDataList)
            {
                var filePath = Path.Combine(folderPath, $"table_{tableData.Index}.html");

                using (var writer = new StreamWriter(filePath))
                {
                    writer.Write(tableData.TableData);
                }
            }
        }

        private int CountImagesInDocument(WordprocessingDocument document)
        {
            int imageCount = 0;

            foreach (var part in document.MainDocumentPart.ImageParts)
            {
                imageCount++;
            }

            return imageCount;
        }

        private List<string> GetImageBase64List(WordprocessingDocument document)
        {
            var imageBase64List = new List<string>();

            foreach (var part in document.MainDocumentPart.ImageParts)
            {
                using (var stream = part.GetStream())
                using (var memoryStream = new MemoryStream())
                {
                    stream.CopyTo(memoryStream);
                    var imageBytes = memoryStream.ToArray();
                    var imageBase64 = Convert.ToBase64String(imageBytes);
                    imageBase64List.Add(imageBase64);
                }
            }

            return imageBase64List;
        }

        private void SaveImages(WordprocessingDocument document, string folderPath)
        {
            int imageCount = 0;

            foreach (var part in document.MainDocumentPart.ImageParts)
            {
                using (var stream = part.GetStream())
                {
                    var imagePath = Path.Combine(folderPath, $"image_{++imageCount}.png");
                    using (var fileStream = new FileStream(imagePath, FileMode.Create))
                    {
                        stream.CopyTo(fileStream);
                    }
                }
            }
        }

        private void SaveTextToFile(string text, string filePath)
        {
            using (var writer = new StreamWriter(filePath))
            {
                writer.WriteLine(text);
            }
        }

        [HttpPost]
        public ActionResult CountImages(HttpPostedFileBase file)
        {
            if (file != null && file.ContentLength > 0)
            {
                try
                {
                    using (Stream stream = file.InputStream)
                    using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(stream, false))
                    {
                        int imageCount = CountImagesInDocument(wordDocument);
                        ViewBag.ImageCount = imageCount;

                        if (imageCount > 0)
                        {
                            string tempFolderPath = Server.MapPath("~/TempFiles");
                            string imagesFolderPath = Path.Combine(tempFolderPath, "images");
                            string tablesFolderPath = Path.Combine(tempFolderPath, "Tables");
                            string textFilePath = Path.Combine(tempFolderPath, "text.txt");

                            // Create temporary folders if they don't exist
                            Directory.CreateDirectory(imagesFolderPath);
                            Directory.CreateDirectory(tablesFolderPath);

                            // Save images to the "images" folder
                            SaveImages(wordDocument, imagesFolderPath);

                            // Save tables to the "Tables" folder
                            var tableDataList = GetTableDataList(wordDocument);
                            SaveTablesToFolder(tableDataList, tablesFolderPath);

                            // Save text to the "text.txt" file
                            SaveTextToFile(wordDocument.MainDocumentPart.Document.InnerText, textFilePath);

                            ViewBag.ImagesFolderPath = imagesFolderPath;
                            ViewBag.TablesFolderPath = tablesFolderPath;
                            ViewBag.TextFilePath = textFilePath;

                            // Initialize ViewBag.ImageData
                            ViewBag.ImageData = GetImageBase64List(wordDocument);
                        }

                        int tableCount = CountTablesInDocument(wordDocument);
                        ViewBag.TableCount = tableCount;

                        if (tableCount > 0)
                        {
                            ViewBag.TableDataList = GetTableDataList(wordDocument);
                        }
                    }
                }
                catch (Exception ex)
                {
                    ViewBag.Error = "Error reading the DOCX file: " + ex.Message;
                }
            }
            else
            {
                ViewBag.Error = "Please select a valid DOCX file.";
            }

            return View("Index");
        }
    }
}
