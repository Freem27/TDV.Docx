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
        internal FootNotes(DocxDocument docx) : base(docx,"w:footnotes")
        {
            docxDocument = docx;
            try
            {
                file = docx.sourceFolder.FindFile("footnotes.xml");
                xmlDoc = new XmlDocument();
                xmlDoc.LoadXml(file.GetSourceString());
                nsmgr = new XmlNamespaceManager(xmlDoc.NameTable);
                nsmgr.AddNamespace("w", xmlDoc.DocumentElement.NamespaceURI);
                xmlEl = (XmlElement)xmlDoc.SelectSingleNode("/w:footnotes", nsmgr);
            }
            catch (FileNotFoundException)
            {
                IsExist = false;
            }
            
        }


        public override string ToString()
        {
            return string.Join(" ",FindChilds<Footnote>().Where(x=>x.Type==FOOTER_TYPE.NONE));
        }
    }

    public enum FOOTER_TYPE { NONE,SEPARATOR, CONTINUATION_SEPAPRATOR }


    public class Footer : BaseNode
    {
        public Relationship Relationship;
        /// <summary>
        /// 
        /// </summary>
        /// <param name="file"></param>
        /// <param name="create">инициализировать новый .xml футера</param>
        internal Footer(DocxDocument docx, ArchFile file, Relationship relationship, bool create = false) : base(docx, "w:ftr")
        {
            this.Relationship = relationship;
            this.file = file;
            xmlDoc = new XmlDocument();
            if(create)
            {
                xmlDoc.LoadXml($@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<w:ftr xmlns:wpc=""http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"" xmlns:mc=""http://schemas.openxmlformats.org/markup-compatibility/2006"" xmlns:o=""urn:schemas-microsoft-com:office:office"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"" xmlns:m=""http://schemas.openxmlformats.org/officeDocument/2006/math"" xmlns:v=""urn:schemas-microsoft-com:vml"" xmlns:wp14=""http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"" xmlns:wp=""http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"" xmlns:w10=""urn:schemas-microsoft-com:office:word"" xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"" xmlns:w14=""http://schemas.microsoft.com/office/word/2010/wordml"" xmlns:w15=""http://schemas.microsoft.com/office/word/2012/wordml"" xmlns:wpg=""http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"" xmlns:wpi=""http://schemas.microsoft.com/office/word/2010/wordprocessingInk"" xmlns:wne=""http://schemas.microsoft.com/office/word/2006/wordml"" xmlns:wps=""http://schemas.microsoft.com/office/word/2010/wordprocessingShape"" mc:Ignorable=""w14 w15 wp14"">
	<w:p >
		<w:pPr>
		</w:pPr>
	</w:p>
</w:ftr>
");
            }
            else
            { 
                xmlDoc.LoadXml(file.GetSourceString());
            }
            nsmgr = new XmlNamespaceManager(xmlDoc.NameTable);
            nsmgr.AddNamespace("w", xmlDoc.DocumentElement.NamespaceURI);
            xmlEl = (XmlElement)xmlDoc.SelectSingleNode("/w:ftr", nsmgr);
            
        }



        public void ComparePageNumbers(DOC_PART_GALLERY_VALUE pageNumbers,HORIZONTAL_ALIGN hAlign=HORIZONTAL_ALIGN.CENTER,string author="TDV")
        {
            if (pageNumbers == DOC_PART_GALLERY_VALUE.NONE && this.PageNumbers == DOC_PART_GALLERY_VALUE.NONE)
                return;
            else if(this.PageNumbers!= pageNumbers || PageNumbersHorizontalAlign != hAlign)
            {
                this.PageNumbers = pageNumbers;
                if(pageNumbers!=DOC_PART_GALLERY_VALUE.NONE)
                    PageNumbersHorizontalAlign = hAlign;
                CustomXmlInsRangeStart customXmlInsRangeStart = FindChild<CustomXmlInsRangeStart>();
                if(customXmlInsRangeStart==null)
                {
                    customXmlInsRangeStart = NewNodeBefore<CustomXmlInsRangeStart>(Std);
                    customXmlInsRangeStart.Author = author;
                    CustomXmlInsRangeEnd customXmlInsRangeEnd = NewNodeAfter<CustomXmlInsRangeEnd>(Std);
                    customXmlInsRangeEnd.Id = customXmlInsRangeStart.Id;
                }
                Paragraph p = Std.SdtContent.P;
                Ins ins=p.NewNodeLast<Ins>();
                ins.Author = author;
                foreach (Node n in p.childNodes)
                    if (n is R)
                        n.MoveTo(ins);
            }
        }

        public DOC_PART_GALLERY_VALUE PageNumbers
        {
            get
            {
                DocPartGallery docPartGallery= FindChild<Sdt>()?.FindChild<StdPr>()?.FindChild<DocPartObj>()?.FindChild<DocPartGallery>();
                if (docPartGallery == null)
                    return DOC_PART_GALLERY_VALUE.NONE;
                return docPartGallery.Value;
            }
            set
            {
                //убрать pageNumbers, если они в обычном параграфе, а не в блоке SDT
                foreach(Paragraph p in FindChilds<Paragraph>())
                {
                    if (p.Text.Contains("PAGE") && p.Text.Contains("MERGEFORMAT"))
                        p.Delete();
                }
                switch(value)
                {
                    case DOC_PART_GALLERY_VALUE.NONE:
                        if (PageNumbers != DOC_PART_GALLERY_VALUE.NONE)
                            Std.Delete();
                        break;
                    case DOC_PART_GALLERY_VALUE.PAGE_NUMBERS_BOTTOM_OF_PAGE:
                        Std.StdPr.DocPartObj.DocPartGallery.Value = value;
                        Std.StdPr.DocPartObj.DocPartUnique = true;
                        Paragraph p = Std.SdtContent.P;
                        while(p.childNodes.Count>0)
                            p.childNodes.First().Delete();
                        p.pPr.HorizontalAlign = HORIZONTAL_ALIGN.CENTER;
                        R r1 = p.NewNodeLast<R>();
                        r1.NewNodeLast<FldChar>().FldCharType = FLD_CHAR_TYPE.BEGIN;
                        R r2 = p.NewNodeLast<R>();
                        r2.NewNodeLast<InstrText>().Text = "PAGE \\* MERGEFORMAT";
                        R r3 = p.NewNodeLast<R>();
                        r3.NewNodeLast<FldChar>().FldCharType = FLD_CHAR_TYPE.SEPARATE;
                        R r4 = p.NewNodeLast<R>();
                        r4.rPr.NoProof = true;
                        r4.t.Text = "2";
                        R r5 = p.NewNodeLast<R>();
                        r5.NewNodeLast<FldChar>().FldCharType = FLD_CHAR_TYPE.END;
                        break;
                    default:
                        throw new NotImplementedException();
                }
            }
        }
        public HORIZONTAL_ALIGN PageNumbersHorizontalAlign
        {
            get
            {
                return Std.SdtContent.P.pPr.HorizontalAlign;
            }
            set
            {
                Std.SdtContent.P.pPr.HorizontalAlign =value;
            }
        }

        public Sdt Std
        {
            get { return FindChildOrCreate<Sdt>(INSERT_POS.FIRST); }
        }

        public new string Text()
        {
            string result = string.Join(" ", childNodes.Where(x => x is Paragraph).Select(x => ((Paragraph)x).Text));
            return result;
        }

        public override void ApplyAllFixes()
        {
            foreach (Node n in childNodes)
            {
                if (n is Paragraph)
                {
                    Paragraph p = (Paragraph)n;
                    p.ApplyAllFixes();
                }
                else if (n is Table)
                {
                    Table t = (Table)n;
                    t.ApplyAllFixes();
                }
                else if (n is SectProp)
                {
                    n.FindChild<SectPrChange>()?.Delete();
                }
                else if (n is CustomXmlInsRangeStart)
                    n.Delete();
                else if (n is CustomXmlInsRangeEnd)
                    n.Delete();
                else if (n is Sdt)
                    ((Sdt)n).ApplyAllFixes();
            }
        }

    }

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

    public class CustomXmlInsRangeStart : Node
    {
        public CustomXmlInsRangeStart() : base("w:customXmlInsRangeStart") { }
        public CustomXmlInsRangeStart(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:customXmlInsRangeStart") {
            
        }
        public int Id
        {
            get
            {
                try { 
                    return Int32.Parse(xmlEl.GetAttribute("id",xmlEl.NamespaceURI));
                }
                catch
                {
                    return 0;
                }
            }
            set
            {
                xmlEl.SetAttribute("id", xmlEl.NamespaceURI, (value).ToString());
            }
        }
        public string Author
        {
            get
            {
                return xmlEl.GetAttribute("w:author");
            }
            set
            {
                xmlEl.SetAttribute("author", xmlEl.NamespaceURI, value);
            }
        }

        public DateTime? Date
        {
            get
            {
                try
                {
                    return DateTime.Parse(xmlEl.GetAttribute("w:date"));
                }catch
                {
                    return null;
                }
            }
            set
            {
                if(value==null)
                    xmlEl.RemoveAttribute("date", xmlEl.NamespaceURI);
                else
                    xmlEl.SetAttribute("date", xmlEl.NamespaceURI, ((DateTime)value).ToString("yyyy-MM-ddTHH:mm:ssZ"));
            }
        }

        public override void InitXmlElement()
        {
            base.InitXmlElement();
            if (string.IsNullOrEmpty(xmlEl.GetAttribute("id", xmlEl.NamespaceURI)))
                xmlEl.SetAttribute("id", xmlEl.NamespaceURI, (xmlDoc.GetLastId(0) + 1).ToString());
            Author = "TDV";
            xmlEl.SetAttribute("date", xmlEl.NamespaceURI, DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ssZ"));
        }

        public override string ToString()
        {
            return $"w:customXmlInsRangeStart id={Id} Author={Author}";
        }
    }


    /// <summary>
    /// Закрывающий тег для  w:customXmlInsRangeStart
    /// !!!! Важно id должен быть равен соответсвующему открывающему тегу customXmlInsRangeStart.id
    /// </summary>
    public class CustomXmlInsRangeEnd : Node
    {
        public CustomXmlInsRangeEnd() : base("w:customXmlInsRangeEnd") { }
        public CustomXmlInsRangeEnd(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:customXmlInsRangeEnd")
        {
        }
        public int Id
        {
            get
            {
                try
                {
                    return Int32.Parse(xmlEl.GetAttribute("w:id", xmlEl.NamespaceURI));
                }
                catch
                {
                    return 0;
                }
            }
            set
            {
                xmlEl.SetAttribute("id", xmlEl.NamespaceURI, (value).ToString());
            }
        }

        public override string ToString()
        {
            return $"w:customXmlInsRangeEnd id={Id}";
        }

    }

    public class Sdt : Node
    {
        public Sdt() : base("w:sdt") { }
        public Sdt(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:sdt") { }
        public StdPr StdPr
        {
            get
            {
                return FindChildOrCreate<StdPr>(INSERT_POS.FIRST);
            }
        }
        public SdtContent SdtContent
        {
            get
            {
                return FindChildOrCreate<SdtContent>();
            }
        }

        public override void ApplyAllFixes()
        {
            foreach (Node n in childNodes)
            {
                if (n is Paragraph)
                {
                    ((Paragraph)n).ApplyAllFixes();
                }
                else if (n is Table)
                {
                    ((Table)n).ApplyAllFixes();
                }
                else if (n is SectProp)
                {
                    n.FindChild<SectPrChange>()?.Delete();
                }
                else if (n is CustomXmlInsRangeStart)
                    n.Delete();
                else if (n is CustomXmlInsRangeEnd)
                    n.Delete();
                else if (n is SdtContent)
                    ((SdtContent)n).ApplyAllFixes();
            }
        }
    }

    public class StdPr : Node
    {
        public StdPr() : base("w:sdtPr") { }
        public StdPr(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:sdtPr") {
        }
        public DocPartObj DocPartObj
        {
            get
            {
                return FindChildOrCreate<DocPartObj>();
            }
        }
        public override void InitXmlElement()
        {
            base.InitXmlElement();
            NewNodeFirst<IdNode>().Value = GenerateId(9).ToString();
        }
    }

    public class IdNode : Node
    {
        public IdNode() : base("w:id") { }
        public IdNode(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:id") { }
        public string Value
        {
            get
            {
                return xmlEl.GetAttribute("w:val");
            }
            set
            {
                xmlEl.SetAttribute("val", nsmgr.LookupNamespace("w"), value);
            }
        }
    }

    public class SdtContent : Node
    {
        public SdtContent() : base("w:sdtContent") { }
        public SdtContent(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:sdtContent") { }
        public Paragraph P
        {
            get
            {
                return FindChildOrCreate<Paragraph>();
            }
        }
    }

    public class DocPartObj : Node
    {
        public DocPartObj() : base("w:docPartObj") { }
        public DocPartObj(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:docPartObj") { }
        public DocPartGallery DocPartGallery
        {
            get {
                return FindChildOrCreate<DocPartGallery>(INSERT_POS.FIRST);
            } 
        }

        public bool DocPartUnique
        {
            get
            {
                bool result = childNodes.Where(x => x.xmlEl.Name == "docPartUnique").Any();
                return result;
            }
            set
            {
                if(value)
                    if(!childNodes.Where(x => x.xmlEl.Name == "w:docPartUnique").Any())
                    {
                        xmlEl.AppendChild(xmlDoc.CreateElement("w:docPartUnique", xmlDoc.DocumentElement.NamespaceURI));
                    }
                else
                {
                    XmlElement forDel = childNodes.Where(x => x.xmlEl.Name == "docPartUnique").FirstOrDefault()?.xmlEl;
                    if (forDel != null)
                        xmlDoc.RemoveChild(forDel);
                }
            }
        }
    }
    public enum DOC_PART_GALLERY_VALUE  {NONE, UNKNOWN, PAGE_NUMBERS_BOTTOM_OF_PAGE}
    public class DocPartGallery : Node
    {
        public DocPartGallery() : base("w:docPartGallery") { }
        public DocPartGallery(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:docPartGallery") { }
        public DOC_PART_GALLERY_VALUE Value
        {
            get{
                DOC_PART_GALLERY_VALUE result = DOC_PART_GALLERY_VALUE.UNKNOWN;
                switch (xmlEl.GetAttribute("w:val"))
                {
                    case "Page Numbers (Bottom of Page)":
                        result = DOC_PART_GALLERY_VALUE.PAGE_NUMBERS_BOTTOM_OF_PAGE;
                        break;
                }
                return result;
            
            }
            set
            {
                switch (value)
                {
                    default:
                    case DOC_PART_GALLERY_VALUE.UNKNOWN:
                        xmlEl.RemoveAttribute("val", xmlEl.NamespaceURI);
                        break;
                    case DOC_PART_GALLERY_VALUE.PAGE_NUMBERS_BOTTOM_OF_PAGE:
                        xmlEl.SetAttribute("val", xmlEl.NamespaceURI, "Page Numbers (Bottom of Page)");
                        break;
                }
            }
        }
    }

    public enum FLD_CHAR_TYPE { BEGIN, SEPARATE,END}
    public class FldChar : Node
    {
        public FldChar() : base("w:fldChar") { }
        public FldChar(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:fldChar") { }

        public FLD_CHAR_TYPE FldCharType
        {
            get
            {
                switch (xmlEl.GetAttribute("w:fldCharType"))
                {
                    case "begin":
                        return FLD_CHAR_TYPE.BEGIN;                        
                    case "separate":
                        return FLD_CHAR_TYPE.SEPARATE;
                    case "end":
                        return FLD_CHAR_TYPE.END;
                    default:
                        throw new NotImplementedException($"Не реализовано для w:fldCharType={xmlEl.GetAttribute("w:fldCharType")}");
                }
            }
            set
            {
                switch (value)
                {
                    default:
                    case FLD_CHAR_TYPE.BEGIN:

                        xmlEl.SetAttribute("fldCharType", xmlEl.NamespaceURI, "begin");
                        break;
                    case FLD_CHAR_TYPE.SEPARATE:
                        xmlEl.SetAttribute("fldCharType", xmlEl.NamespaceURI, "separate");
                        break;
                    case FLD_CHAR_TYPE.END:
                        xmlEl.SetAttribute("fldCharType", xmlEl.NamespaceURI, "end");
                        break;
                }
            }
        }
    }
    public class InstrText : Node
    {
        public InstrText() : base("w:instrText") { }
        public InstrText(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:instrText") { }
    }
    public class FooterReference : Node
    {
        public FooterReference() : base("w:footerReference") { }
        public FooterReference(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:footerReference") { }
        public string Id
        {
            get
            {
                return xmlEl.GetAttribute("r:id");
            }
            set
            {
                SetAttribute("r:id", value);
                
            }
        }


        public REFERENCE_TYPE Type
        {
            get
            {
                switch(GetAttribute("w:type"))
                {
                    case "first":
                        return REFERENCE_TYPE.FIRST;
                    case "even":
                        return REFERENCE_TYPE.EVEN;
                    case "default":
                        return REFERENCE_TYPE.DEFAULT;
                }
                throw new NotImplementedException($"Не реализовано для типа {xmlEl.GetAttribute("w:type")}"); 
            }
            set
            {
                string stringType = "";
                switch (value)
                {
                    case REFERENCE_TYPE.FIRST:
                        stringType = "first";
                        break;
                    case REFERENCE_TYPE.EVEN:
                        stringType = "even";
                        break;
                    case REFERENCE_TYPE.DEFAULT:
                        stringType = "default";
                        break;
                }
                SetAttribute("w:type", stringType);
            }
        }

    }
    
}
