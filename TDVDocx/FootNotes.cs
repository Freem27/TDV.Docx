using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace TDV.Docx
{
    public class FootNotes : BaseNode
    {
        private ArchFile file;

        private XmlDocument xmlDoc;
        internal FootNotes(DocxDocument docx) : base("w:footnotes")
        {
            docxDocument = docx;
            file = docx.sourceFolder.FindFile("footnotes.xml"); ;
            xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(file.GetSourceString());
            nsmgr = new XmlNamespaceManager(xmlDoc.NameTable);
            nsmgr.AddNamespace("w", xmlDoc.DocumentElement.NamespaceURI);
            xmlEl = (XmlElement)xmlDoc.SelectSingleNode("/w:footnotes", nsmgr);
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

        public override string ToString()
        {
            return string.Join(" ",FindChilds<Footnote>().Where(x=>x.Type==FOOTER_TYPE.NONE));
        }
    }

    public enum FOOTER_TYPE { NONE,SEPARATOR, CONTINUATION_SEPAPRATOR }

    public class Footnote : Node
    {
        public Footnote() : base("w:footnote") { }
        public Footnote(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:footnote") { }

        public FOOTER_TYPE Type
        {
            get
            {
                FOOTER_TYPE result = FOOTER_TYPE.NONE;
                switch(xmlEl.GetAttribute("w:type"))
                {
                    case "separator":
                        result = FOOTER_TYPE.SEPARATOR;
                        break;
                    case "continuationSeparator":
                        result = FOOTER_TYPE.CONTINUATION_SEPAPRATOR;
                        break;
                    default:
                        Enum.TryParse<FOOTER_TYPE>(xmlEl.GetAttribute("w:type"), true, out result);
                        break;
                }
                return result;
            }
            set
            {
                switch (value)
                {
                    case FOOTER_TYPE.NONE:
                        xmlEl.RemoveAttribute("w:type");
                        break;
                    case FOOTER_TYPE.CONTINUATION_SEPAPRATOR:
                        xmlEl.SetAttribute("type", xmlEl.NamespaceURI, "continuationSeparator");
                        break;
                    default:
                        xmlEl.SetAttribute("type", xmlEl.NamespaceURI, value.ToString().ToLower());
                        break;
                }
            }
        }

        public List<Paragraph> Paragraphs
        {
            get { return FindChilds<Paragraph>(); }
        }

        public override string Text
        {
            get { return string.Join(" ", Paragraphs.Where(x => !x.IsEmpty).Select(x => x.Text)); }
        }

        public override string ToString()
        {
            return Text;
        }
    }

}
