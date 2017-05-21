using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.ServiceModel.Web;
using System.Text;
using Syncfusion.DocToPDFConverter;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;
using Syncfusion.Pdf;
using System.IO;

namespace GeneratePdf_WCFService
{
    // NOTE: You can use the "Rename" command on the "Refactor" menu to change the class name "Service1" in code, svc and config file together.
    // NOTE: In order to launch WCF Test Client for testing this service, please select Service1.svc or Service1.svc.cs at the Solution Explorer and start debugging.
    public class Service1 : IService1
    {
        public string GeneratePdf(string value)
        {
            //Loading word document
            WordDocument document = new WordDocument(new MemoryStream(Convert.FromBase64String(value)), FormatType.Docx);
            DocToPDFConverter conv = new DocToPDFConverter();
            //Converts the word document to pdf
            PdfDocument pdf = conv.ConvertToPDF(document);
            MemoryStream stream = new MemoryStream();
            //Saves the Pdf document to stream
            pdf.Save(stream);
            //Sets the stream position
            stream.Position = 0;
            pdf.Close();
            document.Close();
            //Returns the stream data as Base 64 string
            return Convert.ToBase64String(stream.ToArray());
        }
    }
}
