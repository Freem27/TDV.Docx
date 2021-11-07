﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.XPath;
using System.Drawing;

namespace TDV.Docx
{
    public class BaseNode : Node
    {
        internal ArchFile file;
        protected BaseNode(DocxDocument docxDocument,string qualifiedName = ""):base(qualifiedName)
        {
            this.docxDocument = docxDocument;
            IsExist = true;
            GetDocxDocument().FilesForApply.Add(this);
        }

        public new virtual void ApplyAllFixes() {
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
                else if (n is CustomXmlInsRangeStart)
                    n.Delete();
                else if (n is CustomXmlInsRangeEnd)
                    n.Delete();
                else if (n is Sdt)
                {
                    ((Sdt)n).ApplyAllFixes();
                }
            }
        }

        internal void FillNamespaces()
        {
            nsmgr = new XmlNamespaceManager(xmlDoc.NameTable);
            IDictionary<string, string> localNamespaces = null;
            XPathNavigator xNav = xmlDoc.CreateNavigator();
            while (xNav.MoveToFollowing(XPathNodeType.Element))
            {
                localNamespaces = xNav.GetNamespacesInScope(XmlNamespaceScope.Local);
                foreach (var localNamespace in localNamespaces)
                {
                    string prefix = localNamespace.Key;
                    if (string.IsNullOrEmpty(prefix))
                        prefix = "DEFAULT";
                    nsmgr.AddNamespace(prefix, localNamespace.Value);
                }
            }
        }

        public void Apply()
        {
            if (!IsExist)
                throw new Exception("numbering.xml does not exist :(");
            using (StringWriter stringWriter = new StringWriter())
            using (XmlWriter xw = XmlWriter.Create(stringWriter))
            {
                xmlDoc.WriteTo(xw);
                xw.Flush();
                file.content = Encoding.UTF8.GetBytes(stringWriter.GetStringBuilder().ToString());
            }
        }
        public bool IsExist;
        public DocxDocument docxDocument;

        public override string ToString()
        {
            return $"Base Node {this.GetType().Name} File={file.Name}";
        }
    }


    public class Section
    {
        public Section(int Pos)
        {
            this.Pos = Pos;
            childNodes = new List<Node>();
        }
        /// <summary>
        /// Переменная для хранения своих комментариев к секции. 
        /// Не используется внутри библиотеки
        /// </summary>
        public object Tag;
        public List<Node> childNodes;
        public SectProp sectProp;
        public List<T> FindChilds<T>() where T : Node
        {
            return childNodes.Where(x => x is T).Select(x => (T)x).ToList();
        }

        /// <summary>
        /// Порядковый номер секции
        /// </summary>
        public int Pos;
    }

    public class Document : BaseNode
    {
        public Body body
        {
            get { return (Body)childNodes.Where(x => x is Body).FirstOrDefault(); }
        }

        public void UpdateSections()
        {
            _sections = new List<Section>();
            int sectNum = 0;
            Section s = new Section(sectNum);
            foreach (Node n in body.childNodes)
            {
                s.childNodes.Add(n);
                //n.Section = s;
                if (n is Paragraph && n.FindChild<PProp>()?.FindChild<SectProp>() != null)
                {
                    s.sectProp = n.FindChild<PProp>()?.FindChild<SectProp>();
                    _sections.Add(s);
                    sectNum++; 
                    s = new Section(sectNum);
                }else if(n is SectProp)
                {
                    s.sectProp = (SectProp)n;
                    _sections.Add(s);
                    sectNum++;
                    s = new Section(sectNum);
                }
            }
            if(s.childNodes.Count>0)
            {
                _sections.Add(s);
            }
        }

        private List<Section> _sections;

        /// <summary>
        /// Документ может быть разбит на секции, например к разным секциям относятся страницы имеющие разный формат\ориентацию
        /// Метод UpdateSections() обновляет список секций. при первом обращении выполняется автоматически
        /// </summary>
        /// <returns></returns>
        public List<Section> Sections
        {
            get {
                if (_sections == null)
                    UpdateSections();
                return _sections;
            }
        }
        public Document(DocxDocument docx) : base(docx,"w:documnent")
        {
            qualifiedName = "w:document";
            try
            {
                file = docx.sourceFolder.FindFile("document.xml");

                xmlDoc = new XmlDocument();
                xmlDoc.LoadXml(file.GetSourceString());
                FillNamespaces();
                xmlEl = (XmlElement)xmlDoc.SelectSingleNode("/w:document", nsmgr);

                this.xmlDoc = xmlEl.OwnerDocument;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        /// <summary>
        /// Принять все правки
        /// </summary>
        public override void ApplyAllFixes()
        {
            foreach(Node n in body.childNodes)
            {
                if(n is Paragraph)
                {
                    Paragraph p = (Paragraph)n;
                    p.ApplyAllFixes();
                }
                else if(n is Table)
                {
                    Table t = (Table)n;
                    t.ApplyAllFixes();
                }
                else if(n is SectProp)
                {
                    n.FindChild<SectPrChange>()?.Delete();
                }
            }
        }


        /*
        public void Apply()
        {
            RemoveEmptyNodes(xmlEl);
            using (StringWriter stringWriter = new StringWriter())
            using (XmlWriter xw = XmlWriter.Create(stringWriter))
            {
                xmlDoc.WriteTo(xw);
                xw.Flush();
                file.content = Encoding.UTF8.GetBytes(stringWriter.GetStringBuilder().ToString());
            }
        }*/

        /// <summary>
        /// удаляет ноды, в которых не заполенны атрибуты (согласно списку)
        /// </summary>
        /// <param name="el"></param>
        private void RemoveEmptyNodes(XmlElement el)
        {

            List<string> namesForDel = new List<string>() { "w:vAlign", "w:tcBorders", "w:tblBorders", "w:trHeight" };
            List<XmlElement> forDel=new List<XmlElement>();
            foreach (XmlNode item in el.ChildNodes)
            {
                if(item is XmlElement)
                {
                    XmlElement child = (XmlElement)item;
                    if (namesForDel.IndexOf(child.Name) > -1 && child.ChildNodes.Count==0 && child.Attributes.Count==0)
                        forDel.Add(child);
                    else if (child.ChildNodes.Count > 0)
                        RemoveEmptyNodes(child);
                }
            }

            foreach (XmlElement item in forDel)
            {
                el.RemoveChild(item);
            }

        }
    }

    public class Body : Node
    {
        public Body(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:body")
        { }
        public SectProp sectProp
        {
            get
            {
                SectProp result = childNodes.Where(x => x is SectProp).Select(x => (SectProp)x).FirstOrDefault();
                if (result == null)
                    result = new SectProp(this);
                return result;
            }
        }
        
    }

    public class RProp : Node
    {
        public RProp() : base("w:rPr") { }
        public RProp(Node parent) : base(parent, "w:rPr") { }
        public RProp(XmlElement xmlElement, Node parent, string qualifiedName = "w:rPr") : base(xmlElement, parent, qualifiedName) { }


        public void SetStyle(RStyle style)
        {
            IsBold = style.isBold;
            IsItalic = style.isItalic;
            IsStrike = style.isStrike;
            Font = style.font;
            FontSize = style.fontSize;
            Color = style.color;
            Highlight = style.highlight;
            Underline=style.underline;
            border.border = style.border;
        }

        /// <summary>
        /// устанавливает режим правки
        /// mode = "del", "ins"
        /// </summary>
        public void SetCorrectionMode(string mode, string author = "TDV")
        {
            var cNode = (XmlElement)xmlEl.SelectSingleNode($"w:{mode}", nsmgr);
            if (cNode == null)
            {
                cNode = (XmlElement)xmlDoc.CreateElement($"w:{mode}", xmlEl.NamespaceURI);
                cNode.SetAttribute("id", xmlEl.NamespaceURI, (xmlDoc.GetLastId() + 1).ToString());
                xmlEl.InsertBefore(cNode, xmlEl.FirstChild);
            }
            cNode.SetAttribute("author", xmlEl.NamespaceURI, author);
            cNode.SetAttribute("date", xmlEl.NamespaceURI, DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ssZ"));

        }


        public RBorder border
        {
            get
            {
                RBorder result = childNodes.Where(x => x is RBorder).Select(x => (RBorder)x).FirstOrDefault();
                if (result == null)
                    result = new RBorder(this);
                return result;
            }
        }

        public void CompareBorder(Border b, string author = "TDV")
        {
            border.CompareBorder(b, author);
        }

        public void CompareStyle(RStyle style, string author = "TDV")
        {
            CompareBold(style.isBold, author);
            CompareBorder(style.border, author);
            CompareItalic(style.isItalic, author);
            CompareStrike(style.isStrike, author);
            CompareUnderline(style.underline, author);
            CompareColor(style.color, author);
            CompareHighlight(style.highlight, author);
            CompareFontSize(style.fontSize, author);
            CompareFont(style.font, author);
        }

        public void CompareFont(string font, string author = "TDV")
        {
            if (font.ToUpper() != Font.ToUpper())
            {
                CreateChangeNode("w:rPrChange", xmlEl, author);
                Font = font;
            }
        }
        public void CompareFontSize(double fontSize, string author = "TDV")
        {
            if (fontSize != FontSize)
            {
                CreateChangeNode("w:rPrChange", xmlEl, author);
                FontSize = fontSize;
            }
        }
        public void CompareBold(bool isBold, string author = "TDV")
        {
            if (IsBold != isBold)
            {
                CreateChangeNode("w:rPrChange", xmlEl, author);
                IsBold = isBold;
            }
        }
        public void CompareItalic(bool isItalic, string author = "TDV")
        {
            if (IsItalic != isItalic)
            {
                CreateChangeNode("w:rPrChange", xmlEl, author);
                IsItalic = isItalic;
            }
        }
        /// <summary>
        /// Зачеркивание
        /// </summary>
        /// <param name="isStrike"></param>
        /// <param name="author"></param>
        public void CompareStrike(bool isStrike, string author = "TDV")
        {
            if (IsStrike != isStrike)
            {
                CreateChangeNode("w:rPrChange", xmlEl, author);
                IsStrike = isStrike;
            }
        }
        /// <summary>
        /// подчеркивание
        /// </summary>
        /// <param name="underline"></param>
        /// <param name="author"></param>
        public void CompareUnderline(LINE_TYPE underline, string author = "TDV")
        {
            if (Underline != underline)
            {
                CreateChangeNode("w:rPrChange", xmlEl, author);
                Underline = underline;
            }
        }
        /// <summary>
        /// Цвет выделения
        /// </summary>
        /// <param name="highlight"></param>
        /// <param name="author"></param>
        public void CompareHighlight(string highlight, string author = "TDV")
        {
            if (Highlight != highlight)
            {
                CreateChangeNode("w:rPrChange", xmlEl, author);
                Highlight = highlight;
            }
        }
        public void CompareColor(string color, string author = "TDV")
        {
            if (!(string.IsNullOrEmpty(Color) && color=="black") && Color != color)
            {
                CreateChangeNode("w:rPrChange", xmlEl, author);
                Color = color;
            }
        }

        public string Font
        {
            get
            {
                XmlElement fontNode = (XmlElement)xmlEl.SelectSingleNode("w:rFonts", nsmgr);
                if (fontNode != null)
                {
                    if (fontNode.HasAttribute("w:ascii"))
                        return fontNode.GetAttribute("w:ascii");
                    if (fontNode.HasAttribute("w:hAnsi"))
                        return fontNode.GetAttribute("w:hAnsi");
                    if (fontNode.HasAttribute("w:ascii"))
                        return fontNode.GetAttribute("w:cs");
                }
                if (style != null)
                {
                    RProp styleRProp = style.GetStyleProp<RProp>();
                    if (styleRProp != null)
                        return styleRProp.Font;
                }
                return "";
            }
            set
            {
                XmlElement fontNode = (XmlElement)xmlEl.SelectSingleNode("w:rFonts", nsmgr);
                if (fontNode == null)
                {
                    fontNode = xmlDoc.CreateElement("w", "rFonts", xmlDoc.DocumentElement.NamespaceURI);
                    xmlEl.AppendChild(fontNode);
                }
                //var attributes = xmlEl.SelectSingleNode("w:rFonts", nsmgr).Attributes;
                fontNode.SetAttribute("ascii", xmlEl.NamespaceURI, value);
                fontNode.SetAttribute("hAnsi", xmlEl.NamespaceURI, value);
                fontNode.SetAttribute("cs", xmlEl.NamespaceURI, value);
            }

        }

        /// <summary>
        /// без проверки правописания (не проверять орфографию и грамматику)
        /// </summary>
        public bool NoProof
        {
            get
            {
                bool result = childNodes.Where(x => x.xmlEl.Name == "noProof").Any();
                return result;
            }
            set
            {
                if (value)
                    if (!childNodes.Where(x => x.xmlEl.Name == "noProof").Any())
                    {
                        xmlEl.AppendChild(xmlDoc.CreateElement("w:noProof", xmlDoc.DocumentElement.NamespaceURI));
                    }
                    else
                    {
                        XmlElement forDel = childNodes.Where(x => x.xmlEl.Name == "noProof").FirstOrDefault()?.xmlEl;
                        if (forDel != null)
                            xmlDoc.RemoveChild(forDel);
                    }
            }
        }


        public Style style
        {
            get
            {
                XmlElement n = (XmlElement)xmlEl.SelectSingleNode("w:rStyle", nsmgr);
                if (n != null)
                {
                    string styleId = null;
                    styleId = n.GetAttribute("w:val");
                    if (!string.IsNullOrEmpty(styleId))
                    {
                        DocxDocument docx = GetDocxDocument();
                        return docx.styles.GetStyleById(styleId);
                    }
                }
                return null;
            }
            /*set
            {
                XmlElement fontNode = (XmlElement)xmlEl.SelectSingleNode("w:rFonts", nsmgr);
                if (fontNode == null)
                {
                    fontNode = doc.CreateElement("w", "rFonts", doc.DocumentElement.NamespaceURI);
                    xmlEl.AppendChild(fontNode);
                }
                //var attributes = xmlEl.SelectSingleNode("w:rFonts", nsmgr).Attributes;
                fontNode.SetAttribute("ascii", xmlEl.NamespaceURI, value);
                fontNode.SetAttribute("hAnsi", xmlEl.NamespaceURI, value);
                fontNode.SetAttribute("cs", xmlEl.NamespaceURI, value);
            }*/
        }

        public double? FontSize
        {
            get
            {
                XmlElement szEl = (XmlElement)xmlEl.SelectSingleNode("w:sz", nsmgr);
                XmlElement szCsEl = (XmlElement)xmlEl.SelectSingleNode("w:szCs", nsmgr);
                if ((szEl ?? szCsEl) != null)
                {
                    var attributes = (szEl ?? szCsEl).Attributes;
                    return double.Parse(attributes["w:val"].Value) / 2;
                }
                if (style != null)
                {
                    RProp styleRProp = style.GetStyleProp<RProp>();
                    if (styleRProp != null)
                        return styleRProp.FontSize;
                }
                return null;
            }
            set
            {
                XmlElement szEl = (XmlElement)xmlEl.SelectSingleNode("w:sz", nsmgr);
                XmlElement szCsEl = (XmlElement)xmlEl.SelectSingleNode("w:szCs", nsmgr);
                if (szEl == null)
                {
                    szEl = xmlDoc.CreateElement("w", "sz", xmlEl.NamespaceURI);
                    xmlEl.AppendChild(szEl);
                }
                if (szCsEl == null)
                {
                    szCsEl = xmlDoc.CreateElement("w", "szCs", xmlEl.NamespaceURI);
                    xmlEl.AppendChild(szCsEl);
                }
                szEl.SetAttribute("val", xmlEl.NamespaceURI, (value * 2).ToString());
                szCsEl.SetAttribute("val", xmlEl.NamespaceURI, (value * 2).ToString());
            }
        }

        public bool IsBold
        {
            get
            {
                var b = xmlEl.SelectSingleNode("w:b", nsmgr);
                if (b != null)
                    return true;
                if (style != null)
                {
                    RProp styleRProp = style.GetStyleProp<RProp>();
                    if (styleRProp != null)
                        return styleRProp.IsBold;
                }
                return false;
            }
            set
            {
                XmlElement bNode = (XmlElement)xmlEl.SelectSingleNode("w:b", nsmgr);
                if (value == false && bNode != null)
                    xmlEl.RemoveChild(bNode);
                if (value == true && bNode == null)
                {
                    bNode = xmlDoc.CreateElement("w", "b", xmlEl.NamespaceURI);
                    xmlEl.AppendChild(bNode);
                }
            }
        }

        public bool IsItalic
        {
            get
            {
                var b = xmlEl.SelectSingleNode("w:i", nsmgr);
                if (b != null)
                    return true;
                if (style != null)
                {
                    RProp styleRProp = style.GetStyleProp<RProp>();
                    if (styleRProp != null)
                        return styleRProp.IsItalic;
                }
                return false;
            }
            set
            {
                XmlElement iNode = (XmlElement)xmlEl.SelectSingleNode("w:i", nsmgr);
                if (value == false && iNode != null)
                {
                    xmlEl.RemoveChild(iNode);
                }

                if (value == true && iNode == null)
                {
                    iNode = xmlDoc.CreateElement("w", "i", xmlEl.NamespaceURI);
                    xmlEl.AppendChild(iNode);
                }
            }
        }

        /// <summary>
        /// Зачеркнутый
        /// </summary>
        public bool IsStrike
        {
            get
            {
                var n = xmlEl.SelectSingleNode("w:strike", nsmgr);
                if (n != null)
                    return true;
                if (style != null)
                {
                    RProp styleRProp = style.GetStyleProp<RProp>();
                    if (styleRProp != null)
                        return styleRProp.IsStrike;
                }
                return false;
            }
            set
            {
                XmlElement n = (XmlElement)xmlEl.SelectSingleNode("w:strike", nsmgr);
                if (value == false && n != null)
                {
                    xmlEl.RemoveChild(n);
                }

                if (value == true && n == null)
                {
                    n = xmlDoc.CreateElement("w", "strike", xmlEl.NamespaceURI);
                    xmlEl.AppendChild(n);
                }
            }
        }

        public LINE_TYPE Underline
        {
            get
            {
                var n = (XmlElement)xmlEl.SelectSingleNode("w:u", nsmgr);
                if (n != null)
                {
                    if (n.GetAttribute("w:val") != null)
                    {
                        LINE_TYPE result = LINE_TYPE.UNKNOWN;
                        Enum.TryParse<LINE_TYPE>(n.GetAttribute("w:val"), true, out result);
                        return result;
                    }
                }

                if (style != null)
                {
                    RProp styleRProp = style.GetStyleProp<RProp>();
                    if (styleRProp != null)
                        return styleRProp.Underline;
                }
                return LINE_TYPE.NONE;
            }
            set
            {
                XmlElement n = (XmlElement)xmlEl.SelectSingleNode("w:u", nsmgr);
                if (value == LINE_TYPE.NONE && n != null)
                { 
                    xmlEl.RemoveChild(n);
                    return;
                }

                if (value != LINE_TYPE.UNKNOWN)
                {
                    if (n == null)
                        n = xmlDoc.CreateElement("w", "u", xmlEl.NamespaceURI);
                    
                    n.SetAttribute("val", xmlEl.NamespaceURI, value.ToString().ToLower());
                    xmlEl.AppendChild(n);
                }
            }
        }
        /// <summary>
        /// yellow,red,FF0000
        /// </summary>
        public string Highlight
        {
            get
            {
                var n = xmlEl.SelectSingleNode("w:highlight", nsmgr);
                if (n != null && n.Attributes["w:val"] != null)
                    return n.Attributes["w:val"].Value;
                if (style != null)
                {
                    RProp styleRProp = style.GetStyleProp<RProp>();
                    if (styleRProp != null)
                        return styleRProp.Highlight;
                }
                return "";
            }
            set
            {
                XmlElement n = (XmlElement)xmlEl.SelectSingleNode("w:highlight", nsmgr);
                if (String.IsNullOrEmpty(value) && n != null)
                    xmlEl.RemoveChild(n);

                
                if (n == null)
                    n = xmlDoc.CreateElement("w", "highlight", xmlEl.NamespaceURI);

                n.SetAttribute("val", xmlEl.NamespaceURI, string.IsNullOrEmpty(value) ? "auto" : value);
                xmlEl.AppendChild(n);
            }
        }

        /// <summary>
        /// FF0000
        /// </summary>
        public string Color
        {
            get
            {
                var n = xmlEl.SelectSingleNode("w:color", nsmgr);
                if (n != null && n.Attributes["w:val"] != null)
                    return n.Attributes["w:val"].Value;
                
                if (style != null)
                {
                    RProp styleRProp = style.GetStyleProp<RProp>();
                    if (styleRProp != null)
                        return styleRProp.Color;
                }
                return "";
            }
            set
            {
                XmlElement n = (XmlElement)xmlEl.SelectSingleNode("w:color", nsmgr);
                if (String.IsNullOrEmpty(value) && n != null)
                    xmlEl.RemoveChild(n);

                if (n == null)
                    n = xmlDoc.CreateElement("w", "color", xmlEl.NamespaceURI);

                n.SetAttribute("val", xmlEl.NamespaceURI, string.IsNullOrEmpty(value)?"auto":value);
                xmlEl.AppendChild(n);
            }
        }
    }
    public class Highlight : Node
    {
        public Highlight() : base("w:highlight") { }
        public Highlight(Node parent) : base(parent, "w:highlight") { }
        public Highlight(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:highlight") { }
        public string Value
        {
            get
            {
                return GetAttribute("w:val");
            }
            set
            {
                SetAttribute("w:val", value);
            }
        }
    }

        /// <summary>
        /// Отсупы
        /// значения указывать в сантиметрах
        /// </summary>
        public class Ind : Node
    {
        public Ind() : base("w:ind") { }
        public Ind(Node parent) : base(parent, "w:ind") { }
        public Ind(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:ind") { }

        /// <summary>
        /// Отступ первой строки. В сантиметрах
        /// </summary>
        public float firstLine
        {
            get
            {
                if (xmlEl.Attributes["w:firstLine"] != null)
                    return float.Parse(xmlEl.Attributes["w:firstLine"].Value) / 567;
                return 0;
            }
            set
            {
                if (value == -1)
                    return;
                if (value == 0)
                {
                    xmlEl.RemoveAttribute("firstLine", xmlEl.NamespaceURI);
                }
                else
                {
                    xmlEl.RemoveAttribute("hanging", xmlEl.NamespaceURI);
                    xmlEl.SetAttribute("firstLine", xmlEl.NamespaceURI, ((int)(value * 567)).ToString());
                }

            }
        }


        public float left
        {
            get
            {
                if (xmlEl.Attributes["w:left"] != null)
                    return float.Parse(xmlEl.Attributes["w:left"].Value) / 567;
                if (xmlEl.Attributes["w:start"] != null)
                    return float.Parse(xmlEl.Attributes["w:start"].Value) / 567;
                return 0;
            }
            set
            {
                if (value == -1)
                    return;
                if (value == 0)
                {
                    xmlEl.RemoveAttribute("left", xmlEl.NamespaceURI);
                    xmlEl.RemoveAttribute("start", xmlEl.NamespaceURI);
                }
                else
                    xmlEl.SetAttribute("left", xmlEl.NamespaceURI, ((int)(value * 567)).ToString());
            }
        }
        public float right
        {
            get
            {
                if (xmlEl.Attributes["w:right"] != null)
                    return float.Parse(xmlEl.Attributes["w:right"].Value) / 567;
                if (xmlEl.Attributes["w:end"] != null)
                    return float.Parse(xmlEl.Attributes["w:end"].Value) / 567;
                return 0;
            }
            set
            {
                if (value == -1)
                    return;
                if (value == 0)
                {
                    xmlEl.RemoveAttribute("right", xmlEl.NamespaceURI);
                    xmlEl.RemoveAttribute("end", xmlEl.NamespaceURI);
                }
                else
                    xmlEl.SetAttribute("end", xmlEl.NamespaceURI, ((int)(value * 567)).ToString());
            }
        }

        /// <summary>
        /// Отступ первой строки для удаления. Исключает тег firstLine. В сантиметрах
        /// </summary>
        public float hanging
        {
            get
            {
                if (xmlEl.Attributes["w:hanging"] != null)
                    return float.Parse(xmlEl.Attributes["w:hanging"].Value) / 567;
                return 0;
            }
            set
            {
                if (value == -1)
                    return;
                if (value == 0)
                {
                    xmlEl.RemoveAttribute("hanging", xmlEl.NamespaceURI);
                }
                else
                {
                    xmlEl.RemoveAttribute("firstLine", xmlEl.NamespaceURI);
                    xmlEl.SetAttribute("hanging", xmlEl.NamespaceURI, ((int)(value * 567)).ToString());
                }
            }
        }
    }

    public class Spacing : Node
    {
        public Spacing() : base("w:spacing") { }
        public Spacing(Node parent) : base(parent, "w:spacing") { }
        public Spacing(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:spacing") { }

        /// <summary>
        /// Межстрочный интервал.
        /// </summary>
        public float line
        {
            get
            {
                if (xmlEl.Attributes["w:line"] != null)
                    return float.Parse(xmlEl.Attributes["w:line"].Value) / 240;
                return 0;
            }
            set
            {
                if (value != -1)
                    xmlEl.SetAttribute("line", xmlEl.NamespaceURI, ((int)(value * 240)).ToString());
            }
        }

        /// <summary>
        /// Отступ перед абзацем
        /// Значение в пт (как в MS Word)
        /// </summary>
        public float before
        {
            get
            {
                if (xmlEl.Attributes["w:before"] != null)
                    return float.Parse(xmlEl.Attributes["w:before"].Value) / 20;
                return 0;
            }
            set
            {
                if (value != -1)
                    xmlEl.SetAttribute("before", xmlEl.NamespaceURI, ((int)(value * 20)).ToString());
            }
        }

        /// <summary>
        /// Отступ после абзаца
        /// Значение в пт (как в MS Word)
        /// </summary>
        public float after
        {
            get
            {
                if (xmlEl.Attributes["w:after"] != null)
                    return float.Parse(xmlEl.Attributes["w:after"].Value) / 20;
                return 0;
            }
            set
            {
                if (value != -1)
                    xmlEl.SetAttribute("after", xmlEl.NamespaceURI, ((int)(value * 20)).ToString());
            }
        }
    }

    public class PProp : Node
    {
        public PProp() : base("w:pPr") { }

        public PProp(Node parent) : base(parent, "w:pPr") { }

        public PProp(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:pPr") { }

        public void CompareStyle(PStyle style, string author)
        {
            CompareHorizontalAlign(style.horizontalAlign, author);
            CompareInd(style.indentingFirtsLine, style.indentingHanging, style.indentingleft, style.indentingRight,
                author);
            CompareSpacing(style.spacingBefore, style.spacingAfter, style.spacingLine, author);
            CompareBorder(style.borderLeft, style.borderRight, style.borderTop, style.borderBottom, style.borderBetween, style.borderBar, author);
            CompareNumbering(style.numId,style.numLevel, author);
            
        }

        public void SetStyle(PStyle style)
        {
            HorizontalAlign= style.horizontalAlign;
            ind.firstLine = style.indentingFirtsLine;
            ind.hanging = style.indentingHanging;
            ind.left = style.indentingleft;
            ind.right = style.indentingRight;
            spacing.before = style.spacingBefore;
            spacing.after = style.spacingAfter;
            spacing.line = style.spacingLine;
            pBdr.Left = style.borderLeft;
            pBdr.Right = style.borderRight;
            pBdr.Top = style.borderTop;
            pBdr.Bottom = style.borderBottom;
            pBdr.Between = style.borderBetween;
            pBdr.Bar= style.borderBar;
        }

        public bool HasSectPr
        {
            get
            {
                return childNodes.Where(x => x is SectProp).Any();
            }
        }

        public Ind ind
        {
            get
            {
                Ind result = childNodes.Where(x => x is Ind).Select(x => (Ind)x).FirstOrDefault();
                if (result == null)
                    result = new Ind(this);
                return result;
            }
        }
        /// <summary>
        /// Окончание секции. Следующая секция всегда начинается с новой страницы
        /// </summary>
        public SectProp sectPr
        {
            get
            {
                SectProp result = childNodes.Where(x => x is SectProp).Select(x => (SectProp)x).FirstOrDefault();
                if (result == null)
                    result = new SectProp(this);
                return result;
            }
        }

        /// <summary>
        /// Является элементом списка
        /// </summary>
        public bool HasNumPr
        {
            get
            {
                return childNodes.Where(x => x is NumPr).Any();
            }
        }
        public NumPr NumPr
        {
            get
            {
                NumPr result = childNodes.Where(x => x is NumPr).Select(x => (NumPr)x).FirstOrDefault();
                return result;
            }

        }

        /// <summary>
        /// Сравнение оступов. Значения в сантиметрах
        /// </summary>
        /// <param name="firtsLine"></param>
        /// <param name="hanging"></param>
        /// <param name="left"></param>
        /// <param name="right"></param>
        /// <param name="author"></param>
        public void CompareInd(float firtsLine, float hanging = 0, float left = 0, float right = 0, string author = "TDV")
        {
            if (ind.firstLine != firtsLine || ind.hanging != hanging || ind.left != left || ind.right != right)
            {
                CreateChangeNode("w:pPrChange", xmlEl, author);
                ind.firstLine = firtsLine;
                ind.hanging = hanging;
                ind.left = left;
                ind.right = right;
            }
        }


        public void CompareBorder(Border left, Border right, Border top, Border bottom, Border between, Border bar, string author = "TDV")
        {
            pBdr.CompareBorder(BORDER.LEFT, left, author);
            pBdr.CompareBorder(BORDER.RIGHT, right, author);
            pBdr.CompareBorder(BORDER.TOP, top, author);
            pBdr.CompareBorder(BORDER.BOTTOM, bottom, author);
            pBdr.CompareBorder(BORDER.BETWEEN, between, author);
            pBdr.CompareBorder(BORDER.BAR, bar, author);
        }

        /// <summary>
        /// сравнение формата списка
        /// </summary>
        /// <param name="numId">Ссылка на целевой формат</param>
        /// <param name="level">Уровень списка</param>
        /// <param name="author">Автор правки</param>
        public void CompareNumbering(int? numId,int level=0, string author = "TDV")
        {
            if (numId == null || !HasNumPr)
                return;
           
            if(HasNumPr && numId!=NumPr.NumId.Value)
            {
                CreateChangeNode("w:pPrChange", (XmlElement)xmlEl, author);
                NumPr.Level = level;
                NumPr.NumId.Value = (int)numId;
            } else if(!HasNumPr && numId!=null)
            {
                CreateChangeNode("w:pPrChange", (XmlElement)xmlEl, author);
                NewNodeLast<NumPr>();
                NumPr.Level = level;
                NumPr.NumId.Value = (int)numId;
            }
        }

        /// <summary>
        /// отступы
        /// </summary>
        public Spacing spacing
        {
            get
            {
                Spacing result = childNodes.Where(x => x is Spacing).Select(x => (Spacing)x).FirstOrDefault();
                if (result == null)
                    result = new Spacing(this);
                return result;
            }
        }
        public void CompareSpacing(float before = -1, float after = -1, float line = -1, string author = "TDV")
        {
            if (spacing.before != before || spacing.after != after || spacing.line != line)
            {
                CreateChangeNode("w:pPrChange", xmlEl, author);
                spacing.before = before;
                spacing.after = after;
                spacing.line = line;
            }
        }


        public RProp rPr
        {
            get
            {
                RProp result = childNodes.Where(x => x is RProp).Select(x => (RProp)x).FirstOrDefault();
                if (result == null)
                    result = new RProp(this);
                return result;
            }
        }

        public PBorder pBdr
        {
            get
            {
                PBorder result = childNodes.Where(x => x is PBorder).Select(x => (PBorder)x).FirstOrDefault();
                if (result == null)
                    result = new PBorder(this);
                return result;
            }
        }


        public void CompareHorizontalAlign(HORIZONTAL_ALIGN horizontalAlign, string author = "TDV")
        {
            if (HorizontalAlign != horizontalAlign)
            {
                CreateChangeNode("w:pPrChange", xmlEl, author);
                HorizontalAlign = horizontalAlign;
            }
        }

        public HORIZONTAL_ALIGN HorizontalAlign
        {
            get
            {
                Jc jc = FindChild<Jc>();
                if (jc == null)
                    return HORIZONTAL_ALIGN.NONE;
                return jc.Value;
            }
            set
            {
                Jc jc = FindChildOrCreate<Jc>();
                if (value == HORIZONTAL_ALIGN.NONE)
                    jc.Delete();
                else
                    jc.Value = value;
              
            }
        }

        public bool IsBold
        {
            get
            {
                RProp rPr = (RProp)childNodes.Where(x => x is RProp).FirstOrDefault();
                if (rPr != null)
                {
                    if (rPr.IsBold)
                        return true;
                }
                return false;
            }
        }

        public bool IsItalic
        {
            get
            {
                RProp rPr = (RProp)childNodes.Where(x => x is RProp).FirstOrDefault();
                if (rPr != null)
                {
                    if (rPr.IsItalic)
                        return true;
                }
                return false;
            }
        }
    }

    /// <summary>
    /// Горизонтальное выравнивание параграафа
    /// </summary>
    public class Jc : Node
    {
        public Jc() : base("w:jc") { }
        public Jc(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:jc")        { }

        public HORIZONTAL_ALIGN Value
        {
            get
            {
                switch (xmlEl.GetAttribute("w:val"))
                {
                    case "left":
                        return HORIZONTAL_ALIGN.LEFT;
                    case "center":
                        return HORIZONTAL_ALIGN.CENTER;
                    case "right":
                        return HORIZONTAL_ALIGN.RIGHT;
                    case "both":
                        return HORIZONTAL_ALIGN.BOTH;
                    default:
                        return HORIZONTAL_ALIGN.NONE;
                }
            }
            set
            {
                switch (value)
                {
                    case HORIZONTAL_ALIGN.LEFT:
                        xmlEl.SetAttribute("val", xmlEl.NamespaceURI, "left");
                        break;
                    case HORIZONTAL_ALIGN.CENTER:
                        xmlEl.SetAttribute("val", xmlEl.NamespaceURI, "center");
                        break;
                    case HORIZONTAL_ALIGN.RIGHT:
                        xmlEl.SetAttribute("val", xmlEl.NamespaceURI, "right");
                        break;
                    case HORIZONTAL_ALIGN.BOTH:
                        xmlEl.SetAttribute("val", xmlEl.NamespaceURI, "both");
                        break;
                    case HORIZONTAL_ALIGN.NONE:
                        Delete();
                        break;
                    default:
                        break;
                }
            }
        }

        /// <summary>
        /// Синоним для Value
        /// </summary>
        public HORIZONTAL_ALIGN HorizontalAlign
        {
            get { return Value; }
            set { Value = value; }
        }
    }
        public class R : Node
    {
        public R() : base("w:r") { }
        public R(Node parent) : base(parent, "w:r") { }
        public R(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:r") { }

        public RProp rPr
        {
            get
            {
                RProp result = childNodes.Where(x => x is RProp).Select(x => (RProp)x).FirstOrDefault();
                if (result == null)
                {
                    result = new RProp(this);
                }
                return result;
            }
        }

        /// <summary>
        /// Рецензирование - удаление
        /// создает тег w:del  и помещает в него эту ноду
        /// </summary>
        public void CorrectDel(string author = "TDV")
        {
            Del delNode = parent.FindChild<Del>();
            if (delNode == null)
                delNode = parent.NewNodeLast<Del>();
            delNode.Author = author;
            XmlElement del = xmlDoc.CreateElement("w","delText", xmlEl.NamespaceURI);
            del.InnerText = Text;
            xmlEl.AppendChild(del);
            XmlElement tForDel = (XmlElement)xmlEl.SelectSingleNode("w:t", nsmgr);
            if(tForDel!=null)
                xmlEl.RemoveChild(tForDel);
            this.MoveTo(delNode);
        }

        public void CorrectSetText(string newText, string author = "TDV")
        {
            if (Text != newText)
            {
                Ins newIns = parent.NewNodeAfter<Ins>(xmlEl);
                newIns.Author = author;
                R newRnode = newIns.NewNodeLast<R>();
                newRnode.Text = newText;
                CorrectDel(author);
            }
        }

        public override string ToString()
        {
            if (FindChild<Drawing>() != null)
                return "[img]";
            return base.Text;
        }

        public int DrawingCount()
        {
            return FindChildsRecurcieve<Drawing>().Count();
        }

        public new string Text
        {
            get
            {
                return base.Text;
            }
            set
            {
                t.xmlEl.InnerText = value;
            }
        }

        public T t
        {
            get
            {
                return FindChildOrCreate<T>();
            }
        }

        public override void InitXmlElement()
        {
            base.InitXmlElement();
        }

        public bool IsBold
        {
            get {return rPr.IsBold;}
        }

        public bool IsItalic
        {
            get { return rPr.IsItalic; }
        }
    }

    public class T : Node
    {
        public T() : base("w:t") { }
        public T(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:t")
        { }

        public new string Text
        {
            get
            {
                return base.Text;
            }
            set
            {
                if (xmlEl != null)
                    xmlEl.InnerText = value;
            }
        }
    }

    /// <summary>
    /// Контейнер для хранения подсвеченного текста
    /// </summary>
    public class HighlightText
    {
        public HighlightText(Paragraph parentP, int pos, string text="",string color=null)
        {
            Parent = parentP;
            this.Text = text;
            this.Pos = pos;
            this.Color = color;
        }
        public readonly Paragraph Parent;
        public string Text;
        public int Pos;
        public string Color;
    }

    public class Paragraph : Node
    {
        public Paragraph() : base("w:p") { }
        public Paragraph(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:p")
        { }

        public Size Height
        {
            get
            {
                if (Section?.sectProp?.WorkspaceWidth == null)
                    return null;

                //вычислить количество полных строк

                //вычислить высоту каждой из строк

                //перемножить
                throw new NotImplementedException();
            }
        }
        public override string ToString()
        {
            string result = $"";
            foreach(R r in rNodes)
            {
                if (r.DrawingCount() > 0)
                    result += "[img]";
                else
                    result += r.Text;
            }    
            return result;
        }

        public List<HighlightText> GetHighiltText()
        {
            List<HighlightText> result = new List<HighlightText>();
            int pos = 0;
            HighlightText highlightText = new HighlightText(this, pos);
            for(int rIndex=0;rIndex<rNodes.Count();rIndex++)
            {
                R r = rNodes[rIndex];
                Highlight currHighlight = r.FindChild<RProp>()?.FindChild<Highlight>();

                if(currHighlight != null)
                {
                    if (highlightText.Color == null)
                        highlightText.Color = currHighlight.Value;
                    //если подсветка следующей ноды такая же, как у текузей - зааппендить текст. иначе - доавбить в результат и создать новый HighlightText
                    if (rIndex + 1 <= rNodes.Count() - 1 && rNodes[rIndex+1].FindChild<RProp>()?.FindChild<Highlight>()?.Value==currHighlight.Value)
                    {
                        highlightText.Text += r.Text;
                    }else
                    {
                        highlightText.Text += r.Text;
                        result.Add(highlightText);
                        pos++;
                        highlightText = new HighlightText(this, pos);
                    }
                }
            }
            return result;
        }

        public override void ApplyAllFixes()
        {
            //удалить Del ноды
            List<Del> delList = FindChilds<Del>();
            foreach (Del d in delList)
                d.Delete();

            //Применить все ins ноды
            List<Ins> insList = FindChilds<Ins>();
            foreach (Ins ins in insList)
            {
                foreach (Node insNode in ins.childNodes)
                    insNode.MoveTo(this);
                ins.Delete();
            }
            FindChild<PProp>()?.FindChild<PprChange>()?.Delete();
            FindChild<PProp>()?.FindChild<RProp>()?.FindChild<RprChange>()?.Delete();
            FindChild<PProp>()?.FindChild<RProp>()?.FindChild<Ins>()?.Delete();
            FindChild<PProp>()?.FindChild<SectProp>()?.FindChild<SectPrChange>()?.Delete();
            foreach (R r in rNodes)
            { 
                r.FindChild<RProp>()?.FindChild<RprChange>()?.Delete();
            }
        }

        /// <summary>
        /// Возвращает True если в параграфе нет текста и нет изображений
        /// </summary>
        public bool IsEmpty
        {
            get
            {
                if (DrawingCount() == 0 && string.IsNullOrEmpty(Text.Trim()))
                    return true;
                return false;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="style"></param>
        /// <param name="author"></param>
        public void CompareRStyle(RStyle style, string author = "TDV",bool applyToRnodes=true)
        {
            if(applyToRnodes)
            { 
                foreach (R r in rNodes)
                {
                    r.rPr.CompareStyle(style, author);
                }
            }
            pPr.rPr.CompareStyle(style, author);
        }

        public void ComparePStyle(PStyle style, string author = "TDV")
        {
            pPr.CompareStyle(style, author);
        }
        public void CompareStyles(PStyle pstyle,RStyle rstyle, string author = "TDV", bool applyToRnodes = true)
        {
            pPr.CompareStyle(pstyle, author);
            CompareRStyle(rstyle, author, applyToRnodes);
        }

        public List<R> rNodes
        {
            get { return childNodes.Where(x => x is R).Select(x => (R)x).ToList(); }
        }

        public PProp pPr
        {
            get
            {
                PProp result = childNodes.Where(x => x is PProp).Select(x => (PProp)x).FirstOrDefault();
                if (result == null)
                {
                    result = new PProp(this);
                }
                return result;
            }
        }

        public int DrawingCount()
        {
            int result = 0;
            foreach (R r in rNodes)
            {
                foreach (Drawing d in r.FindChilds<Drawing>())
                {
                    result++;
                }
            }
            return result;
        }

        /// <summary>
        /// !!!При установке занчения все R ноды кроме первой будут удалены.
        /// </summary>
        public new string Text
        {
            get
            {
                return string.Join("", rNodes.Select(x => x.Text).ToList());
            }
            set
            {
                var rN = rNodes.FirstOrDefault();//тут может быть эксепшн, если нет ни одной r ноды. но это не точно
                if (rN == null)
                    rN=NewNodeLast<R>();
                foreach (var item in rNodes)
                    xmlEl.RemoveChild(item.xmlEl);
                rN.Text = value;
                xmlEl.AppendChild(rN.xmlEl);
            }
        }

        public int WordsCount
        {
            get
            {
                return Text.Split(' ').Where(x=>!string.IsNullOrEmpty(x)).Count();
            }
        }
        
        public void CorrectSetText(string newText,RStyle rStyle, string author = "TDV")
        {
            if (Text != newText)
            {
                foreach (R r in rNodes)
                    r.CorrectDel();
                Ins newIns = NewNodeLast<Ins>();
                newIns.Author = author;
                R newRnode = newIns.NewNodeLast<R>();
                newRnode.rPr.SetStyle(rStyle);
                newRnode.Text = newText;
            }
        }

        /// <summary>
        /// Рецензирование - удаление
        /// создает тег w:del  и помещает в него эту ноду
        /// </summary>
        public void CorrectDel(string author = "TDV")
        {
            pPr.rPr.SetCorrectionMode("del");
            xmlEl.RemoveAttribute("w:rsidRPr");
            xmlEl.SetAttribute("rsidDel",xmlEl.NamespaceURI,xmlEl.GetAttribute("w:rsidR"));
            while (rNodes.Count() > 0)
                rNodes.First().CorrectDel(author);
            List<Paragraph> parList = parent.FindChilds<Paragraph>();
            if (xmlEl == parList.Last().xmlEl)
            {
                //установить у ноды перед этой признак isDel. убирает лишний перенос на новую строку
                for (int i = parList.Count-1; i > 0; i--)
                {
                        if (parList[i].xmlEl != xmlEl) //предыдущая нода
                        {
                            parList[i].pPr.rPr.SetCorrectionMode("del", author);
                        }
                        break;
                }
                Paragraph newP =  parent.NewNodeAfter<Paragraph>(parList.Last().xmlEl);
                newP.xmlEl.SetAttribute("rsidRPr", xmlEl.NamespaceURI, xmlEl.GetAttribute("w:rsidR"));
            }
        }

        public bool AllRunsIsBold
        {
            get
            {
                bool result = true;
                if (rNodes.Count == 0)
                    return false;
                foreach (R run in rNodes)
                    if (!run.IsBold)
                    {
                        result = false;
                        break;
                    }
                return result;
            }
        }

        public bool AllRunsIsItalic
        {
            get
            {

                bool result = true;
                if (rNodes.Count == 0)
                    return false;
                foreach (R run in rNodes)
                    if (!run.IsItalic)
                    {
                        result = false;
                        break;
                    }
                return result;
            }
        }

        public HORIZONTAL_ALIGN HorizontalAlign
        {
            get
            {
                PProp pPropNode = FindChild<PProp>();
                if (pPropNode != null)
                {
                    return pPropNode.HorizontalAlign;
                }
                return HORIZONTAL_ALIGN.NONE;
            }
        }
    }

    public class PageMargin
    {
        public PageMargin(double top , double right , double bottom , double left , double footer , double header , double gutter)
        {
            Top = new Size(top);
            Right =  new Size(right);
            Left =   new Size(left);
            Bottom = new Size(bottom);
            Footer = new Size(footer);
            Header = new Size(header);
            Gutter = new Size(gutter);
        }
        public PageMargin(Size top=null, Size right = null, Size bottom = null, Size left = null, Size footer = null, Size header = null, Size gutter = null)
        {
            Top = top;
            Right = right;
            Left = left;
            Bottom = bottom;
            Footer = footer;
            Header = header;
            Gutter = gutter;
        }
        /// <summary>
        /// Верхняя граница. Значение в сантиметрах.
        /// </summary>
        public Size Top;
        /// <summary>
        /// Правая граница. Значение в сантиметрах.
        /// </summary>
        public Size Right;
        /// <summary>
        /// Нижняя граница. Значение в сантиметрах.
        /// </summary>
        public Size Bottom;
        /// <summary>
        /// Левая граница. Значение в сантиметрах.
        /// </summary>
        public Size Left;
        /// <summary>
        /// расстояние от верхнего края страницы до верхнего края верхнего колонтитула
        /// </summary>
        public Size Header;
        /// <summary>
        /// расстояние от нижнего края страницы до нижнего края нижнего колонтитула
        /// </summary>
        public Size Footer;
        /// <summary>
        /// Дополнительный отступ страницы (для переплета)
        /// </summary>
        public Size Gutter;
    }

    public class PageMarginNode : Node
    {
        public PageMarginNode() : base("w:pgMar") { }
        public PageMarginNode(Node parent) : base(parent, "w:pgMar") { }
        public PageMarginNode(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:pgMar") { }

        /// <summary>
        /// Верхняя граница
        /// </summary>
        public Size Top
        {
            get
            {
                return new Size(Int32.Parse(GetAttribute("w:top")));
            }
            set
            {
                SetAttribute("w:top",value.ValuePoints.ToString());
            }
        }

        /// <summary>
        /// Правая граница
        /// </summary>
        public Size Right
        {
            get
            {
                return new Size(Int32.Parse(GetAttribute("w:right")));
            }
            set
            {
                SetAttribute("w:right", value.ValuePoints.ToString());
            }
        }

        /// <summary>
        /// Нижняя граница
        /// </summary>
        public Size Bottom
        {
            get
            {
                return new Size(Int32.Parse(GetAttribute("w:bottom")));
            }
            set
            {
                SetAttribute("w:bottom", value.ValuePoints.ToString());
            }
        }

        /// <summary>
        /// Левая граница
        /// </summary>
        public Size Left
        {
            get
            {
                return new Size(Int32.Parse(GetAttribute("w:left")));
            }
            set
            {
                SetAttribute("w:left", value.ValuePoints.ToString());
            }
        }

        /// <summary>
        /// расстояние от верхнего края страницы до верхнего края верхнего колонтитула
        /// </summary>
        public Size Header
        {
            get
            {
                try { 
                    return new Size(Int32.Parse(GetAttribute("w:header")));
                }catch
                {
                    return new Size(0);
                }
            }
            set
            {
                SetAttribute("w:header", value.ValuePoints.ToString());
            }
        }

        /// <summary>
        /// расстояние от нижнего края страницы до нижнего края нижнего колонтитула
        /// </summary>
        public Size Footer
        {
            get
            {
                try { 
                    return new Size(Int32.Parse(GetAttribute("w:footer")));
                }catch
                {
                    return new Size(0);
                }
            }
            set
            {
                SetAttribute("w:footer", value.ValuePoints.ToString());
            }
        }
        /// <summary>
        /// Дополнительный отступ страницы (для переплета)
        /// </summary>
        public Size Gutter
        {
            get
            {
                return new Size(Int32.Parse(GetAttribute("w:gutter")));
            }
            set
            {
                SetAttribute("w:gutter", value.ValuePoints.ToString());
            }
        }

    }

    public class Size : IEquatable<Size>
    {
        internal int _value;
        #region operartors
        public static Size operator -(Size a, Size b)
        {
            return new Size(a.ValuePoints - b.ValuePoints);
        }
        public static Size operator -(Size a, int points)
        {
            return new Size(a.ValuePoints - points);
        }
        public static Size operator -(Size a, double cm)
        {
            return new Size(a.ValuePoints - new Size(cm).ValuePoints);
        }
        public static Size operator -(int points, Size b)
        {
            return new Size(points - b.ValuePoints);
        }
        public static Size operator -(double cm, Size b)
        {
            return new Size(new Size(cm).ValuePoints-b.ValuePoints);
        }

        public static Size operator +(Size a, Size b)
        {
            return new Size(a.ValuePoints + b.ValuePoints);
        }
        public static Size operator +(Size a, int points)
        {
            return new Size(a.ValuePoints + points);
        }
        public static Size operator +(int points, Size b)
        {
            return new Size(b.ValuePoints + points);
        }
        public static Size operator +(Size a, double cm)
        {
            return new Size(a.ValuePoints + new Size(cm).ValuePoints);
        }
        public static Size operator +(double cm, Size b)
        {
            return new Size(b.ValuePoints + new Size(cm).ValuePoints);
        }

        public static bool operator <(Size a, Size b)
        {
            return a._value < b._value;
        }
        public static bool operator <=(Size a, Size b)
        {
            return a._value <= b._value;
        }
        public static bool operator >(Size a, Size b)
        {
            return a._value > b._value;
        }
        public static bool operator >=(Size a, Size b)
        {
            return a._value >= b._value;
        }

        public static bool operator ==(Size a, Size b)
        {
            if (a is null)
            {
                if (b is null)
                    return true;
                return false;
            }
            return a.Equals(b);
        }
        public static bool operator !=(Size a, Size b)
        {
            return !(a==b);
        }
        #endregion
        public Size parentSize;
        public TABLE_WIDTH_TYPE SizeType;

        public int Value
        {
            get
            {
                return _value;
            }
        }

        public int PtcValue
        {
            get
            {
                if (SizeType != TABLE_WIDTH_TYPE.PCT)
                    throw new Exception("Тип значения не PTC");
                return _value;
            }
        }
        public Size(int valuePoints)
        {
            _value = valuePoints;
            SizeType = TABLE_WIDTH_TYPE.DXA;
        }

        public Size(int value,TABLE_WIDTH_TYPE type,Size parentSize=null)
        {
            this.parentSize = parentSize;
            SizeType = type;            
            switch(type)
            {
                case TABLE_WIDTH_TYPE.PCT:
                    if (value < 0 || value > 5000)
                        throw new Exception("Значение должно быть в пределах от 0 до 5000 для типа PTC");
                    _value = value;
                    break;
                case TABLE_WIDTH_TYPE.AUTO:
                case TABLE_WIDTH_TYPE.DXA:
                    _value = value;
                    break;
                case TABLE_WIDTH_TYPE.NIL:
                    _value = 0;
                    break;
                default:
                    throw new NotImplementedException();            
            }
                
        }
        

        public Size(double valueCM)
        {
            this.ValueCM = valueCM;
            SizeType = TABLE_WIDTH_TYPE.DXA;
        }
        public int ValuePoints
        {
            get
            {
                switch (SizeType)
                {
                    case TABLE_WIDTH_TYPE.AUTO:
                    case TABLE_WIDTH_TYPE.DXA:
                        return _value;
                    case TABLE_WIDTH_TYPE.PCT:
                        if (parentSize == null)
                            throw new Exception("Для вычисления необходимо указать parentSize");
                        double percents = _value / 50 / 100;
                        return (int)(percents * parentSize._value);
                    case TABLE_WIDTH_TYPE.NIL:
                        return 0;
                    default:
                        throw new NotImplementedException();
                }
            }
            set
            {
                _value = value;
            }
        }
        public double ValueCM
        {
            get
            {
                return Math.Round(ValuePoints / 567d, 1);
            }
            set
            {
                _value = (int)value*567;
            }
        }

        public int CompareTo(Size other)
        {
            if (other is null)
                return -1;
            return this._value - other._value;
        }

        public bool Equals(Size other)
        {
            if (other is null)
                return false;
            return this._value==other._value;
        }

        public override string ToString()
        {
            return $"Points: {ValuePoints}, CM: {ValueCM}";
        }
    }

    public class PgNumType : Node
    {
        public PgNumType() : base("w:pgNumType") { }
        public PgNumType(Node parent) : base(parent, "w:pgNumType") { }
        public PgNumType(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:pgNumType") { }
        public int Start
        {
            get
            {
                try {
                    return Int32.Parse(GetAttribute("w:start"));
                }
                catch
                {
                    return 1;
                }
            }
            set
            {
                if (value >= 1)
                    Delete();
                else
                    SetAttribute("w:start", value.ToString());
            }
        }
    }
        public class SectProp : Node
    {
        public SectProp() : base("w:sectPr") { }
        public SectProp(Node parent) : base(parent, "w:sectPr") { }
        public SectProp(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:sectPr") { }

        public PageMarginNode PgMar
        {
            get
            {
                return FindChildOrCreate<PageMarginNode>();
            }
        }
        /// <summary>
        /// начало нумерции страниц
        /// </summary>
        public int PgNumStart
        {
            get
            {
                return PgNumType?.Start??1;
            }
            set
            {
                if (value <= 1)
                    PgNumType?.RemoveAttribute("w:start");
                else
                { 
                    FindChildOrCreate<PgNumType>().Start = value;
                }
            }
        }


        public PgNumType PgNumType
        {
            get
            {
                return FindChild<PgNumType>();
            }
        }

        /// <summary>
        /// Ширина рабочей области(ширина страницы - отступы)
        /// </summary>
        public Size WorkspaceWidth
        {
            get
            {
                return PgSz.Width - PgMar.Right-PgMar.Left-PgMar.Gutter;
            }
        }
        public Size WorkspaceHeigth
        {
            get
            {
                return PgSz.Heigth - PgMar.Top - PgMar.Bottom;
            }
        }


        /// <summary>
        /// формат сносок
        /// </summary>
        public FootnotePr FootnotePr
        {
            get
            {
                return FindChildOrCreate<FootnotePr>(INSERT_POS.FIRST);
            }
        }

        public void CompareFooter(NUM_FMT fmt, string author = "TDV")
        {
            if (FootnotePr.numFmt.Value != fmt)
            {
                CreateChangeNode("w:sectPrChange",xmlEl, author);
                FootnotePr.numFmt.Value = fmt;
            }
        }

        public void ComparePageStart(int start,string author="TDV")
        {
            if (PgNumStart != start)
            {
                PgNumStart = start;
            }
        }


        public void CompareIsTitlePg(bool value, string author = "TDV")
        {
            if (IsTitlePg != value)
            {
                CreateChangeNode("w:sectPrChange", xmlEl, author);
                IsTitlePg=value;
            }
        }

        public void ComparePageMargin(PageMargin pageMargin, string author = "TDV")
        {
            ComparePageMargin(pageMargin.Top, pageMargin.Bottom, pageMargin.Left, pageMargin.Right, pageMargin.Header, pageMargin.Footer, pageMargin.Gutter, author);
        }

        /// <summary>
        /// Значения в Санмиметрах
        /// </summary>
        /// <param name="top"></param>
        /// <param name="bottom"></param>
        /// <param name="left"></param>
        /// <param name="right"></param>
        /// <param name="header"></param>
        /// <param name="footer"></param>
        /// <param name="gutter"></param>
        /// <param name="author"></param>
        public void ComparePageMargin(Size top = null, Size bottom = null, Size left = null, Size right = null, Size header = null, Size footer = null, Size gutter = null, string author = "TDV")
        {

            if (PgMar.Top != top || PgMar.Bottom != bottom || PgMar.Left != left || PgMar.Right != right || PgMar.Header != header || PgMar.Footer != footer || PgMar.Gutter != gutter)
            {
                CreateChangeNode("w:sectPrChange", xmlEl, author);
                if(top!=null)
                    PgMar.Top = top;
                if(bottom!=null)
                    PgMar.Bottom = bottom;
                if(left!=null)
                    PgMar.Left = left;
                if (right != null)
                    PgMar.Right = right;
                if(header!=null)
                    PgMar.Header = header;
                if (footer != null)
                    PgMar.Footer = footer;
                if (gutter != null)
                    PgMar.Gutter = gutter;
            }
        }

        public Header GetHeader(REFERENCE_TYPE type, bool createIfNotExist = false)
        {
            string stringType = "unknown";
            switch (type)
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

            XmlElement header = (XmlElement)xmlEl.SelectSingleNode("w:headerReference[@w:type=\"" + stringType + "\"] ", nsmgr);
            /*if (header != null)
            {
                string id=header.GetAttribute("r:id");
                DocxDocument docx = GetDocxDocument();
                ArchFile headerFile= docx.wordRels.GetFileById(id);
                return new Header(docx,headerFile);
            }
            throw new KeyNotFoundException("Не удалось найти файл заголовка");
            */


            DocxDocument docx = GetDocxDocument();
            if (header != null)
            {
                string id = header.GetAttribute("r:id");

                ArchFile footerFile = docx.wordRels.GetFileById(id);
                return new Header(docx, footerFile, docx.wordRels.GetRelationshipById(id));
            }
            else
            {
                if (!createIfNotExist)
                    throw new FileNotFoundException("Не удалось найти файл верхнего колонтитула");
                int maxHeaderIndex = 0;
                ArchFolder wordFolder = docx.sourceFolder.GetFolder("word");
                foreach (ArchFile file in wordFolder.GetFiles())
                {
                    if (file.Name.StartsWith("header"))
                    {
                        int headerIndex = Int32.Parse(file.Name.Replace("header", "").Replace(".xml", ""));
                        if (headerIndex > maxHeaderIndex)
                            maxHeaderIndex = headerIndex;
                    }
                }
                ArchFile newHeaderFile = wordFolder.AddFile($"header{maxHeaderIndex + 1}.xml", new byte[0]);
                Override ov = docx.contentTypes.GetOverride(newHeaderFile.GetFullPath(), true);
                ov.ContentType = Override.ContentTypes.HEADER;
                Relationship newRel = docx.wordRels.NewRelationship(newHeaderFile.Name, RELATIONSIP_TYPE.HEADER);
                Header newHeader = new Header(docx, newHeaderFile, newRel, create: true);
                //прописать в document.xml
                HeaderReference headerReference = docx.document.body.sectProp.GetHeaderReference(type, true);
                headerReference.Id = newRel.Id;

                return newHeader;
            }
        }
         
        /// <summary>
        /// 
        /// </summary>
        /// <param name="type"></param>
        /// <param name="createIfNotExist"> создаст файл футера, если не найдет.пропишет в document.xml.rels. Пропишет в sectProp.</param>
        /// <returns></returns>
        public Footer GetFooter(REFERENCE_TYPE type, bool createIfNotExist = false)
        {
            string stringType = "unknown";
            switch (type)
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


            XmlElement footer = (XmlElement)xmlEl.SelectSingleNode("w:footerReference[@w:type=\"" + stringType + "\"] ", nsmgr);
            DocxDocument docx = GetDocxDocument();
            if (footer != null)
            {
                string id = footer.GetAttribute("r:id");
                
                ArchFile footerFile = docx.wordRels.GetFileById(id);
                return new Footer(docx,footerFile,docx.wordRels.GetRelationshipById(id));
            }else
            {
                if(!createIfNotExist)
                    throw new FileNotFoundException("Не удалось найти файл нижнего колонтитула");
                int maxFooterIndex = 0;
                ArchFolder wordFolder = docx.sourceFolder.GetFolder("word");
                foreach (ArchFile file in wordFolder.GetFiles())
                {
                    if(file.Name.StartsWith("footer"))
                    {
                        int footerIndex = Int32.Parse(file.Name.Replace("footer", "").Replace(".xml", ""));
                        if (footerIndex > maxFooterIndex)
                            maxFooterIndex = footerIndex;
                    }
                }
                ArchFile newFooterFile = wordFolder.AddFile($"footer{maxFooterIndex + 1}.xml", new byte[0]);
                Override ov= docx.contentTypes.GetOverride(newFooterFile.GetFullPath(), true);
                ov.ContentType = Override.ContentTypes.FOOTER;
                Relationship newRel= docx.wordRels.NewRelationship(newFooterFile.Name, RELATIONSIP_TYPE.FOOTER);
                Footer newFooter = new Footer(docx,newFooterFile,newRel, create:true);
                //прописать в document.xml
                FooterReference footerReference= docx.document.body.sectProp.GetFooterReference(type, true);
                footerReference.Id = newRel.Id;
                newFooter.Apply();
                return newFooter;
            }
        }

        public bool IsTitlePg
        {
            get
            {
                bool result = childNodes.Where(x => x.xmlEl.Name == "titlePg").Any();
                return result;
            }
            set
            {
                if (value)
                    if (!childNodes.Where(x => x.xmlEl.Name == "titlePg").Any())
                    {
                        xmlEl.AppendChild(xmlDoc.CreateElement("w:titlePg", xmlDoc.DocumentElement.NamespaceURI));
                    }
                    else
                    {
                        XmlElement forDel = childNodes.Where(x => x.xmlEl.Name == "titlePg").FirstOrDefault()?.xmlEl;
                        if (forDel != null)
                            xmlDoc.RemoveChild(forDel);
                    }
            }
        }

        public FooterReference GetFooterReference(REFERENCE_TYPE type, bool createIfNotExist = false)
        {
            foreach (FooterReference r in FindChilds<FooterReference>())
            {
                if (r.Type == type)
                    return r;
            }
            if(!createIfNotExist)
                throw new KeyNotFoundException();
            FooterReference newFooter = NewNodeFirst<FooterReference>();
            newFooter.Type = type;
            return newFooter;
        }
        public FooterReference GetFooterReference(string id, bool createIfNotExist = false)
        {
            foreach (FooterReference r in FindChilds<FooterReference>())
            {
                if (r.Id==id)
                    return r;
            }
            if (!createIfNotExist)
                throw new KeyNotFoundException();
            FooterReference newFooter = NewNodeFirst<FooterReference>();
            newFooter.Id = id;
            return newFooter;
        }

        public HeaderReference GetHeaderReference(REFERENCE_TYPE type, bool createIfNotExist = false)
        {
            foreach (HeaderReference r in FindChilds<HeaderReference>())
            {
                if (r.Type == type)
                    return r;
            }
            if (!createIfNotExist)
                throw new KeyNotFoundException();
            HeaderReference newHeader = NewNodeFirst<HeaderReference>();
            newHeader.Type = type;
            return newHeader;
        }
        public HeaderReference GetHeaderReference(string id, bool createIfNotExist = false)
        {
            foreach (HeaderReference r in FindChilds<HeaderReference>())
            {
                if (r.Id == id)
                    return r;
            }
            if (!createIfNotExist)
                throw new KeyNotFoundException();
            HeaderReference newHeader = NewNodeFirst<HeaderReference>();
            newHeader.Id = id;
            return newHeader;
        }

        public PgSz PgSz
        {
            get
            {
                return FindChildOrCreate<PgSz>();
            }
        }
    }


    public enum PAGE_ORIENTATION { NONE, PORTRAIT,LANSCAPE}
    /// <summary>
    /// Размер страницы
    /// </summary>
    public class PgSz : Node
    {
        public PgSz() : base("w:pgSz") { }

        public PgSz(Node parent) : base(parent, "w:pgSz") { }
        public PgSz(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:pgSz") { }

        public Size Width
        {
            get
            {
                try
                {
                    return new Size(Int32.Parse(GetAttribute("w:w")));
                }
                catch
                {
                    return new Size(0);
                }
            }
            set
            {
                SetAttribute("w:w", value.ValuePoints.ToString());
            }
        }

        public Size Heigth
        {
            get
            {
                try
                {
                    return new Size (Int32.Parse(GetAttribute("w:h")));
                }
                catch
                {
                    return new Size(0);
                }
            }
            set
            {
                SetAttribute("w:h", value.ValuePoints.ToString());
            }
        }

        public PAGE_ORIENTATION Prientation
        {
            get
            {
                switch (GetAttribute("w:orient"))
                {
                    case "portrait":
                        return PAGE_ORIENTATION.PORTRAIT;
                    case "landscape":
                        return PAGE_ORIENTATION.LANSCAPE;
                    default:
                        return PAGE_ORIENTATION.NONE;
                }
            }
            set
            {
                switch(value)
                {
                    case PAGE_ORIENTATION.LANSCAPE:
                        SetAttribute("w:orient", "landscape");
                        break;
                    case PAGE_ORIENTATION.PORTRAIT:
                        SetAttribute("w:orient", "portrait");
                        break;
                    case PAGE_ORIENTATION.NONE:
                        RemoveAttribute("w:orient");
                        break;
                    default:
                        throw new NotImplementedException();
                }
            }
        }
    }

    public class PBorder : Node
    {
        public PBorder() : base("w:pBdr") { }

        public PBorder(Node parent) : base(parent, "w:pBdr") { }

        public PBorder(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:pBdr") { }

        public void CompareBorder(BORDER type, Border b, string author = "TDV")
        {
            Border currBorder = GetBorder(type);

            if (currBorder != b)
            {
                CreateChangeNode("w:pPrChange", (XmlElement)xmlEl.ParentNode, author);
                SetBorder(type, b);
            }
        }

        public Border Left
        {
            get { return GetBorder(BORDER.BOTTOM); }
            set { SetBorder(BORDER.LEFT, value); }
        }
        public Border Right
        {
            get { return GetBorder(BORDER.BOTTOM); }
            set { SetBorder(BORDER.RIGHT, value); }
        }
        public Border Top
        {
            get { return GetBorder(BORDER.BOTTOM); }
            set { SetBorder(BORDER.TOP, value); }
        }
        public Border Bottom
        {
            get { return GetBorder(BORDER.BOTTOM); }
            set { SetBorder(BORDER.BOTTOM, value); }
        }
        public Border Between
        {
            get { return GetBorder(BORDER.BETWEEN); }
            set { SetBorder(BORDER.BETWEEN, value); }
        }
        public Border Bar
        {
            get { return GetBorder(BORDER.BAR); }
            set { SetBorder(BORDER.BAR, value); }
        }

        private Border GetBorder(BORDER type)
        {
            string prefix = "w";
            string localName = "";
            switch (type)
            {
                case BORDER.LEFT:
                    localName = "left";
                    break;
                case BORDER.RIGHT:
                    localName = "right";
                    break;
                case BORDER.TOP:
                    localName = "top";
                    break;
                case BORDER.BOTTOM:
                    localName = "bottom";
                    break;
                case BORDER.BETWEEN:
                    localName = "between";
                    break;
                case BORDER.BAR:
                    localName = "bar";
                    break;
                default:
                    break;
            }
            XmlElement n = (XmlElement)xmlEl.SelectSingleNode($"{prefix}:{localName}", nsmgr);
            Border b = new Border();
            if (n != null)
            {
                b.color = n.GetAttribute("w:color");
                b.size = Int32.Parse(n.GetAttribute("w:sz"));
                if (n.HasAttribute("w:space"))
                    b.space = Int32.Parse(n.GetAttribute("w:space"));
                else b.space = -1;
                switch (n.GetAttribute("w:val"))
                {
                    case "single":
                        b.type = LINE_TYPE.SINGLE;
                        break;
                    default:
                        b.type = LINE_TYPE.UNKNOWN;
                        break;
                }
            }
            return b;
        }

        private void SetBorder(BORDER type, Border b)
        {
            string prefix = "w";
            string localName = "";
            switch (type)
            {
                case BORDER.LEFT:
                    localName = "left";
                    break;
                case BORDER.RIGHT:
                    localName = "right";
                    break;
                case BORDER.TOP:
                    localName = "top";
                    break;
                case BORDER.BOTTOM:
                    localName = "bottom";
                    break;
                case BORDER.BETWEEN:
                    localName = "between";
                    break;
                case BORDER.BAR:
                    localName = "bar";
                    break;
                default:
                    break;
            }
            XmlElement n = (XmlElement)xmlEl.SelectSingleNode($"{prefix}:{localName}", nsmgr);
            if (b.type == LINE_TYPE.NONE && n != null)
            {
                xmlEl.RemoveChild(n);
            }

            if (b.type != LINE_TYPE.NONE)
            {
                if (n == null)
                {
                    n = xmlDoc.CreateElement($"{prefix}:{localName}", xmlDoc.DocumentElement.NamespaceURI);
                    xmlEl.AppendChild(n);
                }
                switch (b.type)
                {
                    case LINE_TYPE.SINGLE:
                        n.SetAttribute("val", xmlEl.NamespaceURI, "single");
                        break;
                    default:
                        break;
                }
                n.SetAttribute("sz", xmlEl.NamespaceURI, b.size.ToString());
                n.SetAttribute("space", xmlEl.NamespaceURI, b.space.ToString());
                n.SetAttribute("color", xmlEl.NamespaceURI, b.color);
            }
        }

    }

    /// <summary>
    /// Границы Run. нет лево\право\верх\низ. применяется граница по контуру 
    /// </summary>
    public class RBorder : Node
    {
        public RBorder() : base("w:bdr") { }

        public RBorder(Node parent) : base(parent, "w:bdr") { }

        public RBorder(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:bdr") { }

        public void CompareBorder(Border b, string author = "TDV")
        {
            if (b == null)
                return;
            if (border != b)
            {
                CreateChangeNode("w:rPrChange", (XmlElement)xmlEl.ParentNode, author);
                border = b ;
            }
        }

        public Border border
        {
            get
            {
                Border b = new Border();
                if(xmlEl.HasAttribute("w:color"))
                    b.color = xmlEl.GetAttribute("w:color");
                if(xmlEl.HasAttribute("w:sz"))
                    b.size = Int32.Parse(xmlEl.GetAttribute("w:sz"));
                if (xmlEl.HasAttribute("w:space"))
                    b.space = Int32.Parse(xmlEl.GetAttribute("w:space"));
                else b.space = -1;
                switch (xmlEl.GetAttribute("w:val"))
                {
                    case "single":
                        b.type = LINE_TYPE.SINGLE;
                        break;
                    default:
                        b.type = LINE_TYPE.UNKNOWN;
                        break;

                }

                return b;
            }
            set
            {
                if (value.type == LINE_TYPE.NONE)
                {
                    parent.xmlEl.RemoveChild(xmlEl);
                    return;
                }

                else
                {
                    switch ( value.type)
                    {
                        case LINE_TYPE.SINGLE:
                            xmlEl.SetAttribute("val", xmlEl.NamespaceURI, "single");
                            break;
                        default:
                            break;
                    }
                    xmlEl.SetAttribute("sz", xmlEl.NamespaceURI, value.size.ToString());
                    xmlEl.SetAttribute("space", xmlEl.NamespaceURI, value.space.ToString());
                    xmlEl.SetAttribute("color", xmlEl.NamespaceURI, value.color);
                }
            }
        }
        

    }

    public class Border : IEquatable<Border>
    {
        public Border(LINE_TYPE type = LINE_TYPE.NONE, int size = 0, int space = 0, string color = "auto")
        {
            this.type = type;
            this.size = size;
            this.color = color;
            this.space = space;
        }

        public int space;
        public LINE_TYPE type;
        public int size;
        public string color;

        public bool Equals(Border other)
        {
            if (this == null && other == null)
                return true;
            if ((other == null && this != null) || (other != null && this == null))
                return false;
            return size == other.size && space == other.space && type == other.type && color == other.color;
        }
    }
    
    public class Ins : Node
    {

        public Ins() : base("w:ins") { }

        public Ins(Node parent) : base(parent, "w:ins") { }

        public Ins(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:ins") { }

        public string Author
        {
            get { return xmlEl.GetAttribute("author", xmlEl.NamespaceURI); }
            set
            {
                xmlEl.SetAttribute("author", xmlEl.NamespaceURI, value);
            }
        }


        //метод общий для INS и DEL
        public override void InitXmlElement()
        {
            base.InitXmlElement();
            if (string.IsNullOrEmpty(xmlEl.GetAttribute("id", xmlEl.NamespaceURI)))
                xmlEl.SetAttribute("id", xmlEl.NamespaceURI, (xmlDoc.GetLastId() + 1).ToString());
            Author = "TDV";
            xmlEl.SetAttribute("date", xmlEl.NamespaceURI, DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ssZ"));
        }
    }

    public class Del : Node
    {
        public Del() : base("w:del") { }

        public Del(Node parent) : base(parent, "w:del") { }

        public Del(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:del") { }

        public string Author
        {
            get { return xmlEl.GetAttribute("author", xmlEl.NamespaceURI); }
            set
            {
                xmlEl.SetAttribute("author", xmlEl.NamespaceURI, value);
            }
        }


        //метод общий для INS и DEL
        public override void InitXmlElement()
        {
            base.InitXmlElement();
            if (string.IsNullOrEmpty(xmlEl.GetAttribute("id", xmlEl.NamespaceURI)))
                xmlEl.SetAttribute("id", xmlEl.NamespaceURI, (xmlDoc.GetLastId() + 1).ToString());
            Author = "TDV";
            xmlEl.SetAttribute("date", xmlEl.NamespaceURI, DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ssZ"));
        }
    }

    public class Drawing : Node
    {
        public Drawing() : base("w:drawing") { }
        public Drawing(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:drawing") { }

        /// <summary>
        /// только для картинок
        /// </summary>
        /// <returns></returns>
        public Image GetImageFile()
        {
            try
            {
                return FindChild<Inline>()
                    .FindChild<Graphic>().FindChild<GraphicData>().FindChild<Pic>().FindChild<BlipFill>()
                    .FindChild<Blip>()
                    .GetImageFile();
            }
            catch 
            {
                return null;
            }
        }

        /// <summary>
        /// значение в сантиметрах
        /// </summary>
        public float Width
        {
            get
            {
                try
                {
                    return FindChild<Inline>()
                        .FindChild<Extent>().cx/360000.0f;
                }
                catch
                { return -1; }
            }
            set
            {
                try
                {
                    FindChild<Inline>()
                        .FindChild<Extent>().cx = (int)(value * 360000.0f);
                    FindChild<Inline>().FindChild<EffectExtent>().l = 0;
                    FindChild<Inline>().FindChild<EffectExtent>().r = 0;
                    FindChild<Inline>().FindChild<EffectExtent>().b = 0;
                    FindChild<Inline>().FindChild<EffectExtent>().t = 0;
                    FindChild<Inline>()
                        .FindChild<Graphic>().FindChild<GraphicData>().FindChild<Pic>().FindChild<SpPr>().FindChild<Xfrm>().FindChild<Ext>().cx = (int)(value * 360000.0f);
                }
                catch { }
            }
        }
        public float Heigth
        {
            get
            {
                try
                {
                    return FindChild<Inline>()
                        .FindChild<Extent>().cy/ 360000.0f;
                }
                catch
                { return -1; }
            }
            set
            {
                try
                {
                    FindChild<Inline>()
                        .FindChild<Extent>().cy = (int)(value * 360000.0f);
                    FindChild<Inline>().FindChild<EffectExtent>().l = 0;
                    FindChild<Inline>().FindChild<EffectExtent>().r = 0;
                    FindChild<Inline>().FindChild<EffectExtent>().b = 0;
                    FindChild<Inline>().FindChild<EffectExtent>().t = 0;
                    FindChild<Inline>()
                        .FindChild<Graphic>().FindChild<GraphicData>().FindChild<Pic>().FindChild<SpPr>().FindChild<Xfrm>().FindChild<Ext>().cy = (int)(value * 360000.0f);
                }
                catch { }
            }
        }

    }

    public class Inline : Node
    {
        public Inline() : base("wp:inline") { }
        public Inline(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "wp:inline") { }
    }

    public class SectPrChange : Node
    {
        public SectPrChange() : base("wp:sectPrChange") { }
        public SectPrChange(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "wp:sectPrChange") { }
    }

    public class RprChange : Node
    {
        public RprChange() : base("wp:rPrChange") { }
        public RprChange(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "wp:rPrChange") { }
    }
    public class PprChange : Node
    {
        public PprChange() : base("wp:pPrChange") { }
        public PprChange(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "wp:pPrChange") { }
    }

    public class TblPrChange : Node
    {
        public TblPrChange() : base("wp:tblPrChange") { }
        public TblPrChange(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "wp:tblPrChange") { }
    }
    public class TcPrChange : Node
    {
        public TcPrChange() : base("wp:tcPrChange") { }
        public TcPrChange(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "wp:tcPrChange") { }
    }
    public class TblGridChange : Node
    {
        public TblGridChange() : base("wp:tblGridChange") { }
        public TblGridChange(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "wp:tblGridChange") { }
    }
    public class TrPrChange : Node
    {
        public TrPrChange() : base("wp:trPrChange") { }
        public TrPrChange(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "wp:trPrChange") { }
    }
    public class Extent : Node
    {
        public Extent() : base("wp:extent") { }
        public Extent(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "wp:extent") { }
        /// <summary>
        /// Ширина
        /// </summary>
        public int cx
        {
            get
            {
                int result = -1;
                Int32.TryParse(xmlEl.GetAttribute("cx"), out result);
                return result;
            }
            set { xmlEl.SetAttribute("cx", value.ToString()); }
        }
        /// <summary>
        /// Высота
        /// </summary>
        public int cy
        {
            get
            {
                int result = -1;
                Int32.TryParse(xmlEl.GetAttribute("cy"), out result);
                return result;
            }
            set{ xmlEl.SetAttribute("cy", value.ToString()); }
        }
    }

    public class EffectExtent : Node
    {
        public EffectExtent() : base("wp:effectExtent") { }
        public EffectExtent(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "wp:effectExtent") { }

        public int l
        {
            get
            {
                int result = -1;
                Int32.TryParse(xmlEl.GetAttribute("l"), out result);
                return result;
            }
            set { xmlEl.SetAttribute("l", value.ToString()); }
        }
        public int r
        {
            get
            {
                int result = -1;
                Int32.TryParse(xmlEl.GetAttribute("r"), out result);
                return result;
            }
            set { xmlEl.SetAttribute("r", value.ToString()); }
        }
        public int t
        {
            get
            {
                int result = -1;
                Int32.TryParse(xmlEl.GetAttribute("t"), out result);
                return result;
            }
            set { xmlEl.SetAttribute("t", value.ToString()); }
        }
        public int b
        {
            get
            {
                int result = -1;
                Int32.TryParse(xmlEl.GetAttribute("b"), out result);
                return result;
            }
            set { xmlEl.SetAttribute("b", value.ToString()); }
        }
    }

    public class SpPr : Node
    {
        public SpPr() : base("pic:spPr") { }
        public SpPr(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "pic:spPr") { }
    }

    public class Xfrm : Node
    {
        public Xfrm() : base("a:xfrm") { }
        public Xfrm(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "a:xfrm") { }
    }

    public class Ext : Node
    {
        public Ext() : base("a:ext") { }
        public Ext(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "a:ext") { }

        /// <summary>
        /// Ширина
        /// </summary>
        public int cx
        {
            get
            {
                int result = -1;
                Int32.TryParse(xmlEl.GetAttribute("cx"), out result);
                return result;
            }
            set { xmlEl.SetAttribute("cx", value.ToString()); }
        }
        /// <summary>
        /// Высота
        /// </summary>
        public int cy
        {
            get
            {
                int result = -1;
                Int32.TryParse(xmlEl.GetAttribute("cy"), out result);
                return result;
            }
            set { xmlEl.SetAttribute("cy", value.ToString()); }
        }
    }
    
    public class Graphic : Node
    {
        public Graphic() : base("a:graphic") { }
        public Graphic(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "a:graphic") { }
    }

    public class GraphicData : Node
    {
        public GraphicData() : base("a:graphicData") { }
        public GraphicData(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "a:graphicData") { }
    }

    public class Pic : Node
    {
        public Pic() : base("pic:pic") { }
        public Pic(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "pic:pic") { }
    }

    public class BlipFill : Node
    {
        public BlipFill() : base("pic:blipFill") { }
        public BlipFill(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "pic:blipFill") { }
    }

    public class Blip : Node
    {
        public Blip() : base("a:blip") { }
        public Blip(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "a:blip") { }

        /// <summary>
        /// 
        /// </summary>
        public string embed
        {
            get { return xmlEl.GetAttribute("r:embed"); }
        }

        public Image GetImageFile()
        {
            byte[] bytes = GetDocxDocument().wordRels.GetFileById(embed).content;
            return Bitmap.FromStream(new MemoryStream(bytes));
            
        }
    }

    public class FootnotePr : Node
    {
        public FootnotePr() : base("w:footnotePr") { }
        public FootnotePr(Node parent) : base(parent, "w:footnotePr") { }
        public FootnotePr(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:footnotePr") { }
        public NumFmt numFmt{
            get
            {
                NumFmt result = FindChild<NumFmt>();
                if (result == null)
                    result = NewNodeLast<NumFmt>();
                return result;
            }
        }
    }

    public enum NUM_FMT { UNKNOWN,
        /// <summary>
        /// 1, 2, 3
        /// </summary>
        DEFAULT,
        /// <summary>
        /// A, B, C
        /// </summary>
        UPPER_LETTER,
        /// <summary>
        /// i, ii, iii
        /// </summary>
        LOWER_ROMAN,
        /// <summary>
        /// I, II, III
        /// </summary>
        UPPER_ROMAN,
        /// <summary>
        /// спецсимволы
        /// </summary>
        CHICAGO,

        BULLET

    }
    /// <summary>
    /// Формат сносок
    /// </summary>
    public class NumFmt : Node
    {
        public NumFmt() : base("w:numFmt") { }
        public NumFmt(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:numFmt") { }

        public NUM_FMT Value
        {
            get
            {
                NUM_FMT result = NUM_FMT.UNKNOWN;

                switch (xmlEl.GetAttribute("w:val"))
                {
                    case "":
                        result=NUM_FMT.DEFAULT;
                        break;
                    case "upperLetter":
                        result = NUM_FMT.UPPER_LETTER;
                        break;
                    case "lowerRoman":
                        result = NUM_FMT.LOWER_ROMAN;
                        break;
                    case "upperRoman":
                        result = NUM_FMT.UPPER_ROMAN;
                        break;
                    case "chicago":
                        result = NUM_FMT.CHICAGO;
                        break;
                    case "bullet":
                        result = NUM_FMT.BULLET;
                        break;
                    default:
                        break;
                }
                return result;
            }
            set
            {
                switch (value)
                {
                    default:
                    case NUM_FMT.DEFAULT:
                        xmlEl.RemoveAttribute("val", xmlEl.NamespaceURI);
                        break;
                    case NUM_FMT.CHICAGO:
                        xmlEl.SetAttribute("val", xmlEl.NamespaceURI, "chicago");
                        break;
                    case NUM_FMT.LOWER_ROMAN:
                        xmlEl.SetAttribute("val", xmlEl.NamespaceURI, "lowerRoman");
                        break;
                    case NUM_FMT.UPPER_ROMAN:
                        xmlEl.SetAttribute("val", xmlEl.NamespaceURI, "upperRoman");
                        break;
                    case NUM_FMT.UPPER_LETTER:
                        xmlEl.SetAttribute("val", xmlEl.NamespaceURI, "upperLetter");
                        break;
                    case NUM_FMT.BULLET:
                        xmlEl.SetAttribute("val", xmlEl.NamespaceURI, "bullet");
                        break;
                }
            }
        }
    }


    /// <summary>
    /// формат списка
    /// </summary>
    /*public enum NUM_FORMAT
    {
        UNKNOWN,
        /// <summary>
        /// Кружки
        /// </summary>
        CIRCLE=1,
        /// <summary>
        /// Треугольники
        /// </summary>
        TRIANGLE=2,
        /// <summary>
        /// О
        /// </summary>
        O=3,
        /// <summary>
        /// квадрат
        /// </summary>
        SQUARE=4,
        /// <summary>
        /// 4 Ромба
        /// </summary>
        FOUR_ROMBUS=5,
        COLOR_ICON=6,
        /// <summary>
        /// черно-белая стрелка
        /// </summary>
        ARROW=7,
        /// <summary>
        /// галочка
        /// </summary>
        CHECK_MARK=8,
        /// <summary>
        /// 1. 2. 3.
        /// </summary>
        NUMBERS1 = 9,
        /// <summary>
        /// 1) 2) 3)
        /// </summary>
        NUMBERS2 = 10,
        /// <summary>
        /// I. II. III.
        /// </summary>
        NUMBERS3 = 11,
        /// <summary>
        /// A. B. C.
        /// </summary>
        SYMBOLS1 = 12,
        /// <summary>
        /// a) b) c)
        /// </summary>
        SYMBOLS2 = 13,
        /// <summary>
        /// a. b. c.
        /// </summary>
        SYMBOLS3 = 14,
        /// <summary>
        /// i. ii. iii.
        /// </summary>
        SYMBOLS4 =15

    }*/

    /// <summary>
    /// Уровень списка
    /// </summary>
    public class Ilvl : Node
    {
        public Ilvl() : base("w:ilvl") { }
        public Ilvl(Node parent) : base(parent, "w:ilvl") {
            Value = 0;
        }
        public Ilvl(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:ilvl") { }
        public int Value
        {
            get
            {
                return Int32.Parse(xmlEl.GetAttribute("w:val"));
            }
            set
            {
                xmlEl.SetAttribute("val", xmlEl.NamespaceURI, value.ToString());
            }
        }
    }

    /// <summary>
    /// тип списка
    /// </summary>
    public class NumId : Node
    {
        public NumId() : base("w:numId") { }
        public NumId(Node parent) : base(parent, "w:numId")
        {
            Value = 0;
        }
        public NumId(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:numId") { }
        public int Value
        {
            get
            {
                return  Int32.Parse(xmlEl.GetAttribute("w:val"));
            }
            set
            {
                xmlEl.SetAttribute("val", xmlEl.NamespaceURI, value.ToString());
            }
        }
    }

    /// <summary>
    /// Формат списка
    /// </summary>
    public class NumPr : Node
    {
        public NumPr() : base("w:numPr") { }
        public NumPr(Node parent,int numId) : base(parent, "w:numPr")
        {
            Level = 0;
            NumId.Value = numId;
        }
        public NumPr(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:numPr") { }


        public Ilvl Ilvl
        {
            get
            {
                Ilvl result = childNodes.Where(x => x is Ilvl).Select(x => (Ilvl)x).FirstOrDefault();
                if (result == null)
                    result = new Ilvl(this);
                return result;
            }
        }
        public NumId NumId
        {
            get
            {
                NumId result = childNodes.Where(x => x is NumId).Select(x => (NumId)x).FirstOrDefault();
                if (result == null)
                    result = new NumId(this);
                return result;
            }
        }
        /// <summary>
        /// уровень списка. начинается с 0
        /// </summary>
        public int Level
        {
            get
            {
                return Ilvl.Value;
            }
            set
            {
                Ilvl.Value = value;
            }       
        }
    }

}
