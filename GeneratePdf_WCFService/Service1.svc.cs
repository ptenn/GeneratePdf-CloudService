﻿using System;
using Syncfusion.DocToPDFConverter;
using Syncfusion.DocIO.DLS;
using Syncfusion.Pdf;
using System.IO;
using System.Net;
using System.Security.Authentication;
using GeneratePdf_WCFService.Jobs;
using log4net;
using Syncfusion.ExcelChartToImageConverter;
using Syncfusion.ExcelToPdfConverter;
using Syncfusion.Presentation;
using Syncfusion.PresentationToPdfConverter;
using Syncfusion.XlsIO;
using FormatType = Syncfusion.DocIO.FormatType;

namespace GeneratePdf_WCFService
{
    // NOTE: You can use the "Rename" command on the "Refactor" menu to change the class name "Service1" in code, svc and config file together.
    // NOTE: In order to launch WCF Test Client for testing this service, please select Service1.svc or Service1.svc.cs at the Solution Explorer and start debugging.
    public class Service1 : IService1
    {
        private static ILog logger = LogManager.GetLogger(typeof(Service1));

        private static readonly String _authKey = "sfzvwaeru8233gxcs3arfz";

        public string GeneratePdf(string value, string type, string authKey)
        {
            if (!_authKey.Equals(authKey))
            {
                throw new AuthenticationException("AuthKey not valid");
            }
            PdfDocument pdf;
            MemoryStream stream;
            switch (type)
            {
                // Word
                case ".doc":
                case ".docx":
                case ".docm":
                case ".dotm":
                case ".rtf":
                case ".dot":
                case ".dotx":
                    // Retrieve the Blob Uri as a String
                    using (WebClient webClient = new WebClient())
                    {
                        try
                        {
                            // Retrive the Blob URI (Document) and convert it into a bytearray.
                            // From there it is Base-64 encoded for sending via Service call. 
                            byte[] bytes = webClient.DownloadData(value);
                            logger.Info($"111 Loading Word Document with value: {value}");
                            //Loading word document
                            WordDocument document = new WordDocument(new MemoryStream(bytes));
                            logger.Info("Created Word Document");
                            DocToPDFConverter conv = new DocToPDFConverter();
                            //Converts the word document to pdf
                            pdf = conv.ConvertToPDF(document);
                            logger.Info("Converted to PDF");
                            stream = new MemoryStream();
                            //Saves the Pdf document to stream
                            pdf.Save(stream);
                            logger.Info("Saved PDF to Stream");
                            //Sets the stream position
                            stream.Position = 0;
                            pdf.Close();
                            document.Close();
                            //Returns the stream data as Base 64 string
                            string returnData = Convert.ToBase64String(stream.ToArray());
                            logger.Info($"Returing converted stream as Base 64 data: {returnData}");

                            return returnData;

                        }
                        catch (Exception ex)
                        {
                            logger.Error(ex);
                            throw;
                        }
                    }
                    // Clean up old files

                // Excel
                case ".xlsx":
                case ".xlsm":
                case ".xlsb":
                case ".xls":
                case ".xltx":
                case ".csv":
                    // Retrieve the Blob Uri as a String
                    using (WebClient webClient = new WebClient())
                    {
                        try
                        {
                            // Retrive the Blob URI (Document) and convert it into a bytearray.
                            // From there it is Base-64 encoded for sending via Service call. 
                            byte[] bytes = webClient.DownloadData(value);

                            ExcelEngine excelEngine = new ExcelEngine();
                            IApplication application = excelEngine.Excel;

                            application.ChartToImageConverter = new ChartToImageConverter();

                            IWorkbook workbook =
                                application.Workbooks.Open(new MemoryStream(bytes), ExcelOpenType.Automatic);

                            ExcelToPdfConverter convert = new ExcelToPdfConverter(workbook);
                            pdf = convert.Convert();
                            stream = new MemoryStream();

                            //Saves the Pdf document to stream
                            pdf.Save(stream);
                            //Sets the stream position
                            stream.Position = 0;
                            pdf.Close();
                            workbook.Close();
                            //Returns the stream data as Base 64 string
                            return Convert.ToBase64String(stream.ToArray());
                        }
                        catch (Exception ex)
                        {
                            logger.Error(ex);
                            throw;
                        }
                    }
                // PowerPoint
                case ".pptx":
                case ".pptm":
                case ".ppt":
                case ".potx":
                case ".ppsx":
                case ".potm":
                case ".pot":
                case ".ppsm":
                case ".pps":
                    // Retrieve the Blob Uri as a String
                    using (WebClient webClient = new WebClient())
                    {
                        try
                        {
                            // Retrive the Blob URI (Document) and convert it into a bytearray.
                            // From there it is Base-64 encoded for sending via Service call. 
                            byte[] bytes = webClient.DownloadData(value);
                            //Opens a PowerPoint Presentation
                            IPresentation presentation = Presentation.Open(new MemoryStream(bytes));

                            //Creates an instance of ChartToImageConverter and assigns it to ChartToImageConverter property of Presentation
                            presentation.ChartToImageConverter = new Syncfusion.OfficeChartToImageConverter.ChartToImageConverter();

                            //Converts the PowerPoint Presentation into PDF document
                            pdf = PresentationToPdfConverter.Convert(presentation);

                            stream = new MemoryStream();
                            //Saves the PDF document
                            pdf.Save(stream);

                            //Sets the stream position
                            stream.Position = 0;
                            pdf.Close();
                            presentation.Close();
                            //Returns the stream data as Base 64 string
                            return Convert.ToBase64String(stream.ToArray());
                        }
                        catch (Exception ex)
                        {
                            logger.Error(ex);
                            throw;
                        }
                    }
            }
            return null;
        }
    }
}
