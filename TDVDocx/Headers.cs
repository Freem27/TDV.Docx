using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace TDV.Docx
{
    public class Header : BaseNode
    {
        private ArchFile file;

        private XmlDocument xmlDoc;
        internal Header(DocxDocument docx, ArchFile file) : base(docx, "w:hdr")
        {
            this.file = file;
            xmlDoc = new XmlDocument();
            
            xmlDoc.LoadXml(file.GetSourceString());
            nsmgr = new XmlNamespaceManager(xmlDoc.NameTable);
            nsmgr.AddNamespace("w", xmlDoc.DocumentElement.NamespaceURI);
            xmlEl = (XmlElement)xmlDoc.SelectSingleNode("/w:hdr", nsmgr);
        }

        public string Text()
        {
            string result = string.Join(" ", childNodes.Where(x => x is Paragraph).Select(x => ((Paragraph)x).Text));
            return result;
        }

        public void Apply()
        {
            using (StringWriter stringWriter = new StringWriter())
            using (XmlWriter xw = XmlWriter.Create(stringWriter))
            {
                xmlDoc.WriteTo(xw);
                xw.Flush();
                file.content = Encoding.UTF8.GetBytes(stringWriter.GetStringBuilder().ToString());
            }
        }
    }
}
