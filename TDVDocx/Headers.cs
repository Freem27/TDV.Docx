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
        public Relationship Relationship;
        internal Header(DocxDocument docx, ArchFile file, Relationship relationship, bool create = false) : base(docx, "w:hdr")
        {
            this.Relationship = relationship;
            this.file = file;
            XmlDoc = new XmlDocument();
            if(create)
            {
                XmlDoc.LoadXml($@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<w:hdr xmlns:wpc=""http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"" xmlns:mc=""http://schemas.openxmlformats.org/markup-compatibility/2006"" xmlns:o=""urn:schemas-microsoft-com:office:office"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"" xmlns:m=""http://schemas.openxmlformats.org/officeDocument/2006/math"" xmlns:v=""urn:schemas-microsoft-com:vml"" xmlns:wp14=""http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"" xmlns:wp=""http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"" xmlns:w10=""urn:schemas-microsoft-com:office:word"" xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"" xmlns:w14=""http://schemas.microsoft.com/office/word/2010/wordml"" xmlns:w15=""http://schemas.microsoft.com/office/word/2012/wordml"" xmlns:wpg=""http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"" xmlns:wpi=""http://schemas.microsoft.com/office/word/2010/wordprocessingInk"" xmlns:wne=""http://schemas.microsoft.com/office/word/2006/wordml"" xmlns:wps=""http://schemas.microsoft.com/office/word/2010/wordprocessingShape"" mc:Ignorable=""w14 w15 wp14"">
	<w:p >
		<w:pPr>
		</w:pPr>
	</w:p>
</w:hdr>
");
            }
            else
            { 
                XmlDoc.LoadXml(file.GetSourceString());
            }
            FillNamespaces();
            XmlEl = (XmlElement)XmlDoc.SelectSingleNode("/w:hdr", Nsmgr);
        }

        public void CompareStyle(ParagraphStyle pStyle, RunStyle rStyle, string author = "TDV")
        {
            foreach (Paragraph p in FindChilds<Paragraph>())
            {
                p.CompareStyles(pStyle, rStyle, author);
            }
        }

        public new string Text()
        {
            string result = string.Join(" ", ChildNodes.Where(x => x is Paragraph).Select(x => ((Paragraph)x).Text));
            return result;
        }

        public void ComparePageNumbers(DOC_PART_GALLERY_VALUE pageNumbers, HORIZONTAL_ALIGN hAlign=HORIZONTAL_ALIGN.CENTER, string author = "TDV")
        {
            if (this.PageNumbers == DOC_PART_GALLERY_VALUE.NONE || PageNumbersHorizontalAlign != hAlign)
            {
                this.PageNumbers = pageNumbers;
                PageNumbersHorizontalAlign = hAlign;
                CustomXmlInsRangeStart customXmlInsRangeStart = FindChild<CustomXmlInsRangeStart>();
                if (customXmlInsRangeStart == null)
                {
                    customXmlInsRangeStart = NewNodeBefore<CustomXmlInsRangeStart>(Sdt);
                    customXmlInsRangeStart.Author = author;
                    CustomXmlInsRangeEnd customXmlInsRangeEnd = NewNodeAfter<CustomXmlInsRangeEnd>(Sdt);
                    customXmlInsRangeEnd.Id = customXmlInsRangeStart.Id;
                }
                Paragraph p = Sdt.SdtContent.P;
                Ins ins = p.NewNodeLast<Ins>();
                ins.Author = author;
                foreach (Node n in p.ChildNodes)
                    if (n is R)
                        n.MoveTo(ins);
            }
        }

        public DOC_PART_GALLERY_VALUE PageNumbers
        {
            get
            {
                DocPartGallery docPartGallery = FindChild<Sdt>()?.FindChild<StdPr>()?.FindChild<DocPartObj>()?.FindChild<DocPartGallery>();
                if (docPartGallery == null)
                    return DOC_PART_GALLERY_VALUE.NONE;
                return docPartGallery.Value;
            }
            set
            {
                switch (value)
                {
                    case DOC_PART_GALLERY_VALUE.NONE:
                        if (PageNumbers != DOC_PART_GALLERY_VALUE.NONE)
                            Sdt.Delete();
                        break;
                    case DOC_PART_GALLERY_VALUE.PAGE_NUMBERS_BOTTOM_OF_PAGE:
                        Sdt.StdPr.DocPartObj.DocPartGallery.Value = value;
                        Sdt.StdPr.DocPartObj.DocPartUnique = true;
                        Paragraph p = Sdt.SdtContent.P;
                        while (p.ChildNodes.Count > 0)
                            p.ChildNodes.First().Delete();
                        p.PProp.HorizontalAlign = HORIZONTAL_ALIGN.CENTER;
                        R r1 = p.NewNodeLast<R>();
                        r1.NewNodeLast<FldChar>().FldCharType = FLD_CHAR_TYPE.BEGIN;
                        R r2 = p.NewNodeLast<R>();
                        r2.NewNodeLast<InstrText>().Text = "PAGE \\* MERGEFORMAT";
                        R r3 = p.NewNodeLast<R>();
                        r3.NewNodeLast<FldChar>().FldCharType = FLD_CHAR_TYPE.SEPARATE;
                        R r4 = p.NewNodeLast<R>();
                        r4.RProp.NoProof = true;
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
                return Sdt.SdtContent.P.PProp.HorizontalAlign;
            }
            set
            {
                Sdt.SdtContent.P.PProp.HorizontalAlign = value;
            }
        }

        public Sdt Sdt
        {
            get { return FindChildOrCreate<Sdt>(INSERT_POS.FIRST); }
        }


        public override void ApplyAllFixes()
        {
            foreach (Node n in ChildNodes)
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

    public class HeaderReference : Node
    {
        public HeaderReference() : base("w:headerReference") { }
        public HeaderReference(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:headerReference") { }
        public string Id
        {
            get
            {
                return XmlEl.GetAttribute("r:id");
            }
            set
            {
                XmlEl.SetAttribute("id", Nsmgr.LookupNamespace("r"), value);
            }
        }

        public Header Header
        {
            get
            {
                if (string.IsNullOrEmpty(Id))
                    return null;
                return GetDocxDocument().GetHeader(Id);
            }
        }
        public REFERENCE_TYPE Type
        {
            get
            {
                return EnumExtentions.ToEnum<REFERENCE_TYPE>(GetAttribute("w:type"));
            }
            set
            {
                SetAttribute("w:type",value.ToStringValue());
            }
        }

    }
}
