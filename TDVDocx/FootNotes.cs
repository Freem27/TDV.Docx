using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace TDV.Docx {
    public class FootNotes : BaseNode {
        internal FootNotes(DocxDocument docx) : base(docx, "w:footnotes") {
            DocxDocument = docx;
            try {
                file = docx.sourceFolder.FindFile("footnotes.xml");
                XmlDoc = new XmlDocument();
                XmlDoc.LoadXml(file.GetSourceString());
                FillNamespaces();
                XmlEl = (XmlElement)XmlDoc.SelectSingleNode("/w:footnotes", Nsmgr);
            }
            catch (FileNotFoundException) {
                IsExist = false;
            }
        }

        public override string ToString() {
            return string.Join(" ", FindChilds<Footnote>().Where(x => x.Type == FOOTER_TYPE.NONE));
        }
    }

    public enum FOOTER_TYPE { NONE, SEPARATOR, CONTINUATION_SEPAPRATOR }

    public class Footer : BaseNode {
        public Relationship Relationship;
        /// <summary>
        /// 
        /// </summary>
        /// <param name="file"></param>
        /// <param name="create">инициализировать новый .xml футера</param>
        internal Footer(DocxDocument docx, ArchFile file, Relationship relationship, bool create = false) : base(docx, "w:ftr") {
            this.Relationship = relationship;
            this.file = file;
            XmlDoc = new XmlDocument();
            if (create) {
                XmlDoc.LoadXml($@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<w:ftr xmlns:wpc=""http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"" xmlns:mc=""http://schemas.openxmlformats.org/markup-compatibility/2006"" xmlns:o=""urn:schemas-microsoft-com:office:office"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"" xmlns:m=""http://schemas.openxmlformats.org/officeDocument/2006/math"" xmlns:v=""urn:schemas-microsoft-com:vml"" xmlns:wp14=""http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"" xmlns:wp=""http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"" xmlns:w10=""urn:schemas-microsoft-com:office:word"" xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"" xmlns:w14=""http://schemas.microsoft.com/office/word/2010/wordml"" xmlns:w15=""http://schemas.microsoft.com/office/word/2012/wordml"" xmlns:wpg=""http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"" xmlns:wpi=""http://schemas.microsoft.com/office/word/2010/wordprocessingInk"" xmlns:wne=""http://schemas.microsoft.com/office/word/2006/wordml"" xmlns:wps=""http://schemas.microsoft.com/office/word/2010/wordprocessingShape"" mc:Ignorable=""w14 w15 wp14"">
	<w:p >
		<w:pPr>
		</w:pPr>
	</w:p>
</w:ftr>
");
            }
            else {
                XmlDoc.LoadXml(file.GetSourceString());
            }
            FillNamespaces();
            XmlEl = (XmlElement)XmlDoc.SelectSingleNode("/w:ftr", Nsmgr);
        }

        public void CompareStyle(ParagraphStyle pStyle, RunStyle rStyle, string author = "TDV") {
            foreach (Paragraph p in FindChilds<Paragraph>()) {
                p.CompareStyles(pStyle, rStyle, author);
            }
        }

        public void ComparePageNumbers(DOC_PART_GALLERY_VALUE pageNumbers, HORIZONTAL_ALIGN hAlign = HORIZONTAL_ALIGN.CENTER, string author = "TDV") {
            if (pageNumbers == DOC_PART_GALLERY_VALUE.NONE && this.PageNumbers == DOC_PART_GALLERY_VALUE.NONE)
                return;
            else if (this.PageNumbers != pageNumbers || PageNumbersHorizontalAlign != hAlign) {
                this.PageNumbers = pageNumbers;
                if (pageNumbers != DOC_PART_GALLERY_VALUE.NONE)
                    PageNumbersHorizontalAlign = hAlign;
                CustomXmlInsRangeStart customXmlInsRangeStart = FindChild<CustomXmlInsRangeStart>();
                if (customXmlInsRangeStart == null) {
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

        public DOC_PART_GALLERY_VALUE PageNumbers {
            get {
                DocPartGallery docPartGallery = FindChild<Sdt>()?.FindChild<StdPr>()?.FindChild<DocPartObj>()?.FindChild<DocPartGallery>();
                if (docPartGallery == null)
                    return DOC_PART_GALLERY_VALUE.NONE;
                return docPartGallery.Value;
            }
            set {
                //убрать pageNumbers, если они в обычном параграфе, а не в блоке SDT
                foreach (Paragraph p in FindChilds<Paragraph>()) {
                    if (p.Text.Contains("PAGE") && p.Text.Contains("MERGEFORMAT"))
                        p.Delete();
                }
                switch (value) {
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
                        InstrText it = r2.NewNodeLast<InstrText>();
                        it.Text = "PAGE \\* MERGEFORMAT";
                        it.XmlSpace = XML_SPACE.PRESERVE;
                        R r3 = p.NewNodeLast<R>();
                        r3.NewNodeLast<FldChar>().FldCharType = FLD_CHAR_TYPE.SEPARATE;
                        R r4 = p.NewNodeLast<R>();
                        r4.RProp.NoProof = true;
                        r4.t.Text = "2";
                        R r5 = p.NewNodeLast<R>();
                        r5.NewNodeLast<FldChar>().FldCharType = FLD_CHAR_TYPE.END;
                        break;
                    default:
                        throw new NotImplementedException();
                }
            }
        }
        public HORIZONTAL_ALIGN PageNumbersHorizontalAlign {
            get {
                return Sdt.SdtContent.P.PProp.HorizontalAlign;
            }
            set {
                Sdt.SdtContent.P.PProp.HorizontalAlign = value;
            }
        }

        public Sdt Sdt {
            get { return FindChildOrCreate<Sdt>(INSERT_POS.FIRST); }
        }

        public new string Text() {
            return string.Join(" ", ChildNodes.Where(x => x is Paragraph).Select(x => ((Paragraph)x).Text));
        }

        public override void ApplyAllFixes() {
            foreach (Node n in ChildNodes) {
                if (n is Paragraph) {
                    Paragraph p = (Paragraph)n;
                    p.ApplyAllFixes();
                }
                else if (n is Table) {
                    Table t = (Table)n;
                    t.ApplyAllFixes();
                }
                else if (n is SectProp) {
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

    public class Footnote : Node {
        public Footnote() : base("w:footnote") { }
        public Footnote(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:footnote") { }

        public FOOTER_TYPE Type {
            get {
                if (!HasAttribute("w:type"))
                    return FOOTER_TYPE.NONE;
                return EnumExtentions.ToEnum<FOOTER_TYPE>(GetAttribute("w:type"));
            }
            set {
                if (value == FOOTER_TYPE.NONE) {
                    RemoveAttribute("w:type");
                    return;
                }
                else
                    SetAttribute("w:type", value.ToStringValue());
            }
        }

        public List<Paragraph> Paragraphs {
            get { return FindChilds<Paragraph>(); }
        }

        public override string Text {
            get { return string.Join(" ", Paragraphs.Where(x => !x.IsEmpty).Select(x => x.Text)); }
        }

        public override string ToString() {
            return Text;
        }
    }

    public class CustomXmlInsRangeStart : Node {
        public CustomXmlInsRangeStart() : base("w:customXmlInsRangeStart") { }
        public CustomXmlInsRangeStart(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:customXmlInsRangeStart") {

        }
        public int Id {
            get {
                try {
                    return Int32.Parse(XmlEl.GetAttribute("id", XmlEl.NamespaceURI));
                }
                catch {
                    return 0;
                }
            }
            set {
                XmlEl.SetAttribute("id", XmlEl.NamespaceURI, (value).ToString());
            }
        }
        public string Author {
            get {
                return XmlEl.GetAttribute("w:author");
            }
            set {
                XmlEl.SetAttribute("author", XmlEl.NamespaceURI, value);
            }
        }

        public DateTime? Date {
            get {
                try {
                    return DateTime.Parse(XmlEl.GetAttribute("w:date"));
                }
                catch {
                    return null;
                }
            }
            set {
                if (value == null)
                    XmlEl.RemoveAttribute("date", XmlEl.NamespaceURI);
                else
                    XmlEl.SetAttribute("date", XmlEl.NamespaceURI, ((DateTime)value).ToString("yyyy-MM-ddTHH:mm:ssZ"));
            }
        }

        public override void InitXmlElement() {
            base.InitXmlElement();
            if (string.IsNullOrEmpty(XmlEl.GetAttribute("id", XmlEl.NamespaceURI)))
                XmlEl.SetAttribute("id", XmlEl.NamespaceURI, (GetDocxDocument().Document.GetNextId()).ToString());
            Author = "TDV";
            XmlEl.SetAttribute("date", XmlEl.NamespaceURI, DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ssZ"));
        }

        public override string ToString() {
            return $"w:customXmlInsRangeStart id={Id} Author={Author}";
        }
    }

    /// <summary>
    /// Закрывающий тег для  w:customXmlInsRangeStart
    /// !!!! Важно id должен быть равен соответсвующему открывающему тегу customXmlInsRangeStart.id
    /// </summary>
    public class CustomXmlInsRangeEnd : Node {
        public CustomXmlInsRangeEnd() : base("w:customXmlInsRangeEnd") { }
        public CustomXmlInsRangeEnd(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:customXmlInsRangeEnd") { }
        public int Id {
            get {
                try {
                    return Int32.Parse(XmlEl.GetAttribute("w:id", XmlEl.NamespaceURI));
                }
                catch {
                    return 0;
                }
            }
            set {
                XmlEl.SetAttribute("id", XmlEl.NamespaceURI, (value).ToString());
            }
        }

        public override string ToString() {
            return $"w:customXmlInsRangeEnd id={Id}";
        }
    }

    public class Sdt : Node {
        public Sdt() : base("w:sdt") { }
        public Sdt(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:sdt") { }
        public StdPr StdPr {
            get {
                return FindChildOrCreate<StdPr>(INSERT_POS.FIRST);
            }
        }
        public SdtContent SdtContent {
            get {
                return FindChildOrCreate<SdtContent>();
            }
        }

        public override void ApplyAllFixes() {
            foreach (Node n in ChildNodes) {
                if (n is Paragraph) {
                    ((Paragraph)n).ApplyAllFixes();
                }
                else if (n is Table) {
                    ((Table)n).ApplyAllFixes();
                }
                else if (n is SectProp) {
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

    public class StdPr : Node {
        public StdPr() : base("w:sdtPr") { }
        public StdPr(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:sdtPr") {
        }
        public DocPartObj DocPartObj {
            get {
                return FindChildOrCreate<DocPartObj>();
            }
        }
        public override void InitXmlElement() {
            base.InitXmlElement();
            NewNodeFirst<IdNode>().Value = GenerateId(9).ToString();
        }
    }

    public class IdNode : Node {
        public IdNode() : base("w:id") { }
        public IdNode(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:id") { }
        public string Value {
            get {
                return XmlEl.GetAttribute("w:val");
            }
            set {
                XmlEl.SetAttribute("val", Nsmgr.LookupNamespace("w"), value);
            }
        }
    }

    public class SdtContent : Node {
        public SdtContent() : base("w:sdtContent") { }
        public SdtContent(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:sdtContent") { }
        public Paragraph P {
            get {
                return FindChildOrCreate<Paragraph>();
            }
        }
    }

    public class DocPartObj : Node {
        public DocPartObj() : base("w:docPartObj") { }
        public DocPartObj(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:docPartObj") { }
        public DocPartGallery DocPartGallery {
            get {
                return FindChildOrCreate<DocPartGallery>(INSERT_POS.FIRST);
            }
        }

        public bool DocPartUnique {
            get {
                bool result = ChildNodes.Where(x => x.XmlEl.Name == "docPartUnique").Any();
                return result;
            }
            set {
                if (value)
                    if (!ChildNodes.Where(x => x.XmlEl.Name == "w:docPartUnique").Any()) {
                        XmlEl.AppendChild(XmlDoc.CreateElement("w:docPartUnique", XmlDoc.DocumentElement.NamespaceURI));
                    }
                    else {
                        XmlElement forDel = ChildNodes.Where(x => x.XmlEl.Name == "docPartUnique").FirstOrDefault()?.XmlEl;
                        if (forDel != null)
                            XmlDoc.RemoveChild(forDel);
                    }
            }
        }
    }
    public enum DOC_PART_GALLERY_VALUE { NONE, PAGE_NUMBERS_BOTTOM_OF_PAGE, PAGE_NUMBERS_TOP_OF_PAGE }
    public class DocPartGallery : Node {
        public DocPartGallery() : base("w:docPartGallery") { }
        public DocPartGallery(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:docPartGallery") { }
        public DOC_PART_GALLERY_VALUE Value {
            get {
                return EnumExtentions.ToEnum<DOC_PART_GALLERY_VALUE>(XmlEl.GetAttribute("w:val"));
            }
            set {
                if (value == DOC_PART_GALLERY_VALUE.NONE) {
                    RemoveAttribute("w:val");
                    return;
                }
                else
                    SetAttribute("w:val", value.ToStringValue());
            }
        }
    }

    public enum FLD_CHAR_TYPE { BEGIN, SEPARATE, END }
    public class FldChar : Node {
        public FldChar() : base("w:fldChar") { }
        public FldChar(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:fldChar") { }

        public FLD_CHAR_TYPE FldCharType {
            get {
                return EnumExtentions.ToEnum<FLD_CHAR_TYPE>(GetAttribute("w:fldCharType"));
            }
            set {
                SetAttribute("w:fldCharType", value.ToStringValue());
            }
        }
    }
    public class InstrText : Node {
        public InstrText() : base("w:instrText") { }
        public InstrText(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:instrText") { }
        public XML_SPACE XmlSpace {
            get {
                if (!HasAttribute("xml:space"))
                    return XML_SPACE.NONE;
                return EnumExtentions.ToEnum<XML_SPACE>(GetAttribute("xml:space"));
            }
            set {
                if (value == XML_SPACE.NONE) {
                    RemoveAttribute("xml:space");
                    return;
                }

                SetAttribute("xml:space", value.ToStringValue());
            }
        }
    }
    public class FooterReference : Node {
        public FooterReference() : base("w:footerReference") { }
        public FooterReference(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:footerReference") { }
        public string Id {
            get {
                return XmlEl.GetAttribute("r:id");
            }
            set {
                SetAttribute("r:id", value);

            }
        }
        public Footer Footer {
            get {
                if (string.IsNullOrEmpty(Id))
                    return null;
                return GetDocxDocument().GetFooter(Id);
            }
        }

        public REFERENCE_TYPE Type {
            get { return EnumExtentions.ToEnum<REFERENCE_TYPE>(GetAttribute("w:type")); }
            set { SetAttribute("w:type", value.ToStringValue()); }
        }
    }
}