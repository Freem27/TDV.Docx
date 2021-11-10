using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.XPath;
using System.Xml.Linq;
using System.Drawing;
using System.Text.RegularExpressions;
using System.Windows.Forms;
 
namespace TDV.Docx
{
    public class BaseNode : Node
    {
        internal ArchFile file;
        protected BaseNode(DocxDocument docxDocument,string qualifiedName = ""):base(qualifiedName)
        {
            this.DocxDocument = docxDocument;
            IsExist = true;
            GetDocxDocument().FilesForApply.Add(this);
        }

        public new virtual void ApplyAllFixes() {
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
            Nsmgr = new XmlNamespaceManager(XmlDoc.NameTable);
            IDictionary<string, string> localNamespaces = null;
            XPathNavigator xNav = XmlDoc.CreateNavigator();
            while (xNav.MoveToFollowing(XPathNodeType.Element))
            {
                localNamespaces = xNav.GetNamespacesInScope(XmlNamespaceScope.Local);
                foreach (var localNamespace in localNamespaces)
                {
                    string prefix = localNamespace.Key;
                    if (string.IsNullOrEmpty(prefix))
                        prefix = "DEFAULT";
                    Nsmgr.AddNamespace(prefix, localNamespace.Value);
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
                XmlDoc.WriteTo(xw);
                xw.Flush();
                file.Content = Encoding.UTF8.GetBytes(stringWriter.GetStringBuilder().ToString());
            }
        }
        public bool IsExist;
        public DocxDocument DocxDocument;

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
            ChildNodes = new List<Node>();
        }
        /// <summary>
        /// Переменная для хранения своих комментариев к секции. 
        /// Не используется внутри библиотеки
        /// </summary>
        public object Tag;
        public List<Node> ChildNodes;
        public SectProp SectProp;
        public List<T> FindChilds<T>() where T : Node
        {
            return ChildNodes.Where(x => x is T).Select(x => (T)x).ToList();
        }

        /// <summary>
        /// Порядковый номер секции
        /// </summary>
        public int Pos;
    }

    public class Document : BaseNode
    {
        public Body Body
        {
            get { return (Body)ChildNodes.Where(x => x is Body).FirstOrDefault(); }
        }

        public int GetLastId()
        {
            int result = 0;
            //XmlNodeList insDelList = XmlEl.SelectNodes("//*[@w:id]", Nsmgr);
            XmlNodeList insDelList = XmlEl.SelectNodes(".//@w:id", Nsmgr);
            if (insDelList.Count > 0)
                result = insDelList.Cast<XmlAttribute>().Max(x => Int32.Parse(x.Value));
                /*foreach (XmlAttribute item in insDelList)
                {
                    int elId = int.Parse(item.Value);
                    if (result < elId)
                    {
                        result = elId;
                    }
                }*/
            return result;
        }
        public void UpdateSections()
        {
            _sections = new List<Section>();
            int sectNum = 0;
            Section s = new Section(sectNum);
            foreach (Node n in Body.ChildNodes)
            {
                s.ChildNodes.Add(n);
                //n.Section = s;
                if (n is Paragraph && n.FindChild<PProp>()?.FindChild<SectProp>() != null)
                {
                    s.SectProp = n.FindChild<PProp>()?.FindChild<SectProp>();
                    _sections.Add(s);
                    sectNum++; 
                    s = new Section(sectNum);
                }else if(n is SectProp)
                {
                    s.SectProp = (SectProp)n;
                    _sections.Add(s);
                    sectNum++;
                    s = new Section(sectNum);
                }
            }
            if(s.ChildNodes.Count>0)
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

                XmlDoc = new XmlDocument();
                XmlDoc.LoadXml(file.GetSourceString());
                FillNamespaces();
                XmlEl = (XmlElement)XmlDoc.SelectSingleNode("/w:document", Nsmgr);

                this.XmlDoc = XmlEl.OwnerDocument;
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
            foreach(Node n in Body.ChildNodes)
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
        public SectProp SectProp
        {
            get
            {
                SectProp result = ChildNodes.Where(x => x is SectProp).Select(x => (SectProp)x).FirstOrDefault();
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

        public Lang Lang
        {
            get { return FindChildOrCreate<Lang>(); }
        }

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
            Border.Border = style.border;
        }

        public RStyle GetRStyle()
        {
            return new RStyle(IsBold, Font, (double)FontSize, IsItalic, IsStrike, Underline, Color, Highlight, Border.Border);            
        }

        /// <summary>
        /// устанавливает режим правки
        /// mode = "del", "ins"
        /// </summary>
        public void SetCorrectionMode(string mode, string author = "TDV")
        {
            var cNode = (XmlElement)XmlEl.SelectSingleNode($"w:{mode}", Nsmgr);
            if (cNode == null)
            {
                cNode = (XmlElement)XmlDoc.CreateElement($"w:{mode}", XmlEl.NamespaceURI);
                cNode.SetAttribute("id", XmlEl.NamespaceURI, (GetDocxDocument().Document.GetLastId() + 1).ToString());
                XmlEl.InsertBefore(cNode, XmlEl.FirstChild);
            }
            cNode.SetAttribute("author", XmlEl.NamespaceURI, author);
            cNode.SetAttribute("date", XmlEl.NamespaceURI, DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ssZ"));

        }


        public RBorder Border
        {
            get
            {
                return FindChildOrCreate<RBorder>();
            }
        }

        public void CompareBorder(Border b, string author = "TDV")
        {
            Border.CompareBorder(b, author);
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
                CreateChangeNode("w:rPrChange", XmlEl, author);
                Font = font;
            }
        }
        public void CompareFontSize(double fontSize, string author = "TDV")
        {
            if (fontSize != FontSize)
            {
                CreateChangeNode("w:rPrChange", XmlEl, author);
                FontSize = fontSize;
            }
        }
        public void CompareBold(bool isBold, string author = "TDV")
        {
            if (IsBold != isBold)
            {
                CreateChangeNode("w:rPrChange", XmlEl, author);
                IsBold = isBold;
            }
        }
        public void CompareItalic(bool isItalic, string author = "TDV")
        {
            if (IsItalic != isItalic)
            {
                CreateChangeNode("w:rPrChange", XmlEl, author);
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
                CreateChangeNode("w:rPrChange", XmlEl, author);
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
                CreateChangeNode("w:rPrChange", XmlEl, author);
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
                CreateChangeNode("w:rPrChange", XmlEl, author);
                Highlight = highlight;
            }
        }
        public void CompareColor(string color, string author = "TDV")
        {
            if (!(string.IsNullOrEmpty(Color) && color=="black") && Color != color)
            {
                CreateChangeNode("w:rPrChange", XmlEl, author);
                Color = color;
            }
        }

        public string Font
        {
            get
            {
                XmlElement fontNode = (XmlElement)XmlEl.SelectSingleNode("w:rFonts", Nsmgr);
                if (fontNode != null)
                {
                    if (fontNode.HasAttribute("w:ascii"))
                        return fontNode.GetAttribute("w:ascii");
                    if (fontNode.HasAttribute("w:hAnsi"))
                        return fontNode.GetAttribute("w:hAnsi");
                    if (fontNode.HasAttribute("w:ascii"))
                        return fontNode.GetAttribute("w:cs");
                }
                if (Style != null)
                {
                    RProp styleRProp = Style.GetStyleProp<RProp>();
                    if (styleRProp != null)
                        return styleRProp.Font;
                }
                return "";
            }
            set
            {
                XmlElement fontNode = (XmlElement)XmlEl.SelectSingleNode("w:rFonts", Nsmgr);
                if (fontNode == null)
                {
                    fontNode = XmlDoc.CreateElement("w", "rFonts", XmlDoc.DocumentElement.NamespaceURI);
                    XmlEl.AppendChild(fontNode);
                }
                //var attributes = xmlEl.SelectSingleNode("w:rFonts", nsmgr).Attributes;
                fontNode.SetAttribute("ascii", XmlEl.NamespaceURI, value);
                fontNode.SetAttribute("hAnsi", XmlEl.NamespaceURI, value);
                fontNode.SetAttribute("cs", XmlEl.NamespaceURI, value);
            }

        }

        /// <summary>
        /// без проверки правописания (не проверять орфографию и грамматику)
        /// </summary>
        public bool NoProof
        {
            get
            {
                return FindChild<NoProof>() != null;
            }
            set
            {
                if (value)
                    FindChildOrCreate<NoProof>();
                else
                    FindChild<NoProof>()?.Delete();
            }
        }


        public Style Style
        {
            get
            {
                return FindChild<RStyleNode>()?.Style??null;
            }
            set
            {
                FindChildOrCreate<RStyleNode>().Style = value;
            }
        }

        public double? FontSize
        {
            get
            {
                Sz sz = FindChild<Sz>();
                SzCs szCs = FindChild<SzCs>();
                if (sz!=null)
                {
                    return sz.Value / 2;
                }else if(szCs != null)
                {
                    return szCs.Value / 2;
                }
                if (Style != null)
                {
                    RProp styleRProp = Style.GetStyleProp<RProp>();
                    if (styleRProp != null)
                        return styleRProp.FontSize;
                }
                return null;
            }
            set
            {
                if(value==null)
                {
                    FindChild<Sz>()?.Delete();
                    FindChild<SzCs>()?.Delete();

                }   else
                { 
                    FindChildOrCreate<Sz>().Value = value * 2;
                    FindChildOrCreate<SzCs>().Value = value * 2;
                }

            }
        }

        public bool IsBold
        {
            get
            {
                var b = XmlEl.SelectSingleNode("w:b", Nsmgr);
                if (b != null)
                    return true;
                if (Style != null)
                {
                    RProp styleRProp = Style.GetStyleProp<RProp>();
                    if (styleRProp != null)
                        return styleRProp.IsBold;
                }
                return false;
            }
            set
            {
                XmlElement bNode = (XmlElement)XmlEl.SelectSingleNode("w:b", Nsmgr);
                if (value == false && bNode != null)
                    XmlEl.RemoveChild(bNode);
                if (value == true && bNode == null)
                {
                    bNode = XmlDoc.CreateElement("w", "b", XmlEl.NamespaceURI);
                    XmlEl.AppendChild(bNode);
                }
            }
        }

        public bool IsItalic
        {
            get
            {
                var b = XmlEl.SelectSingleNode("w:i", Nsmgr);
                if (b != null)
                    return true;
                if (Style != null)
                {
                    RProp styleRProp = Style.GetStyleProp<RProp>();
                    if (styleRProp != null)
                        return styleRProp.IsItalic;
                }
                return false;
            }
            set
            {
                XmlElement iNode = (XmlElement)XmlEl.SelectSingleNode("w:i", Nsmgr);
                if (value == false && iNode != null)
                {
                    XmlEl.RemoveChild(iNode);
                }

                if (value == true && iNode == null)
                {
                    iNode = XmlDoc.CreateElement("w", "i", XmlEl.NamespaceURI);
                    XmlEl.AppendChild(iNode);
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
                var n = XmlEl.SelectSingleNode("w:strike", Nsmgr);
                if (n != null)
                    return true;
                if (Style != null)
                {
                    RProp styleRProp = Style.GetStyleProp<RProp>();
                    if (styleRProp != null)
                        return styleRProp.IsStrike;
                }
                return false;
            }
            set
            {
                XmlElement n = (XmlElement)XmlEl.SelectSingleNode("w:strike", Nsmgr);
                if (value == false && n != null)
                {
                    XmlEl.RemoveChild(n);
                }

                if (value == true && n == null)
                {
                    n = XmlDoc.CreateElement("w", "strike", XmlEl.NamespaceURI);
                    XmlEl.AppendChild(n);
                }
            }
        }

        public LINE_TYPE Underline
        {
            get
            {
                U u = FindChild<U>();
                if (u != null)
                    return u.Value;
                return Style?.FindChild<RProp>().Underline ?? LINE_TYPE.NONE;
            }
            set
            {
                if (value == LINE_TYPE.NONE)
                    FindChild<U>()?.Delete();
                else
                    FindChildOrCreate<U>().Value = value;
            }
        }
        /// <summary>
        /// yellow,red,FF0000
        /// </summary>
        public string Highlight
        {
            get
            {
                var n = XmlEl.SelectSingleNode("w:highlight", Nsmgr);
                if (n != null && n.Attributes["w:val"] != null)
                    return n.Attributes["w:val"].Value;
                if (Style != null)
                {
                    RProp styleRProp = Style.GetStyleProp<RProp>();
                    if (styleRProp != null)
                        return styleRProp.Highlight;
                }
                return "";
            }
            set
            {
                XmlElement n = (XmlElement)XmlEl.SelectSingleNode("w:highlight", Nsmgr);
                if (String.IsNullOrEmpty(value) && n != null)
                    XmlEl.RemoveChild(n);

                
                if (n == null)
                    n = XmlDoc.CreateElement("w", "highlight", XmlEl.NamespaceURI);

                n.SetAttribute("val", XmlEl.NamespaceURI, string.IsNullOrEmpty(value) ? "auto" : value);
                XmlEl.AppendChild(n);
            }
        }

        /// <summary>
        /// FF0000
        /// </summary>
        public string Color
        {
            get
            {
                var n = XmlEl.SelectSingleNode("w:color", Nsmgr);
                if (n != null && n.Attributes["w:val"] != null)
                    return n.Attributes["w:val"].Value;
                
                if (Style != null)
                {
                    RProp styleRProp = Style.GetStyleProp<RProp>();
                    if (styleRProp != null)
                        return styleRProp.Color;
                }
                return "";
            }
            set
            {
                XmlElement n = (XmlElement)XmlEl.SelectSingleNode("w:color", Nsmgr);
                if (String.IsNullOrEmpty(value) && n != null)
                    XmlEl.RemoveChild(n);

                if (n == null)
                    n = XmlDoc.CreateElement("w", "color", XmlEl.NamespaceURI);

                n.SetAttribute("val", XmlEl.NamespaceURI, string.IsNullOrEmpty(value)?"auto":value);
                XmlEl.AppendChild(n);
            }
        }
    }

    public class NoProof : Node
    {
        public NoProof() : base("w:noProof") { }
        public NoProof(Node parent) : base(parent, "w:noProof") { }
        public NoProof(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:noProof") { }
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
    /// 
    /// </summary>
    public class Ind : Node
    {
        public Ind() : base("w:ind") { }
        public Ind(Node parent) : base(parent, "w:ind") { }
        public Ind(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:ind") { }

        /// <summary>
        /// Отступ первой строки. В сантиметрах
        /// </summary>
        public float FirstLine
        {
            get
            {
                if (XmlEl.Attributes["w:firstLine"] != null)
                    return float.Parse(XmlEl.Attributes["w:firstLine"].Value) / 567;
                return 0;
            }
            set
            {
                if (value == -1)
                    return;
                if (value == 0)
                {
                    XmlEl.RemoveAttribute("firstLine", XmlEl.NamespaceURI);
                }
                else
                {
                    XmlEl.RemoveAttribute("hanging", XmlEl.NamespaceURI);
                    XmlEl.SetAttribute("firstLine", XmlEl.NamespaceURI, ((int)(value * 567)).ToString());
                }

            }
        }


        public float Left
        {
            get
            {
                if (XmlEl.Attributes["w:left"] != null)
                    return float.Parse(XmlEl.Attributes["w:left"].Value) / 567;
                if (XmlEl.Attributes["w:start"] != null)
                    return float.Parse(XmlEl.Attributes["w:start"].Value) / 567;
                return 0;
            }
            set
            {
                if (value == -1)
                    return;
                if (value == 0)
                {
                    XmlEl.RemoveAttribute("left", XmlEl.NamespaceURI);
                    XmlEl.RemoveAttribute("start", XmlEl.NamespaceURI);
                }
                else
                    XmlEl.SetAttribute("left", XmlEl.NamespaceURI, ((int)(value * 567)).ToString());
            }
        }
        public float Right
        {
            get
            {
                if (XmlEl.Attributes["w:right"] != null)
                    return float.Parse(XmlEl.Attributes["w:right"].Value) / 567;
                if (XmlEl.Attributes["w:end"] != null)
                    return float.Parse(XmlEl.Attributes["w:end"].Value) / 567;
                return 0;
            }
            set
            {
                if (value == -1)
                    return;
                if (value == 0)
                {
                    XmlEl.RemoveAttribute("right", XmlEl.NamespaceURI);
                    XmlEl.RemoveAttribute("end", XmlEl.NamespaceURI);
                }
                else
                    XmlEl.SetAttribute("end", XmlEl.NamespaceURI, ((int)(value * 567)).ToString());
            }
        }

        /// <summary>
        /// Отступ первой строки для удаления. Исключает тег firstLine. В сантиметрах
        /// </summary>
        public float Hanging
        {
            get
            {
                if (XmlEl.Attributes["w:hanging"] != null)
                    return float.Parse(XmlEl.Attributes["w:hanging"].Value) / 567;
                return 0;
            }
            set
            {
                if (value == -1)
                    return;
                if (value == 0)
                {
                    XmlEl.RemoveAttribute("hanging", XmlEl.NamespaceURI);
                }
                else
                {
                    XmlEl.RemoveAttribute("firstLine", XmlEl.NamespaceURI);
                    XmlEl.SetAttribute("hanging", XmlEl.NamespaceURI, ((int)(value * 567)).ToString());
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
        public float Line
        {
            get
            {
                if (XmlEl.Attributes["w:line"] != null)
                    return float.Parse(XmlEl.Attributes["w:line"].Value) / 240;
                return 0;
            }
            set
            {
                if (value != -1)
                    XmlEl.SetAttribute("line", XmlEl.NamespaceURI, ((int)(value * 240)).ToString());
            }
        }

        /// <summary>
        /// Отступ перед абзацем
        /// Значение в пт (как в MS Word)
        /// </summary>
        public float Before
        {
            get
            {
                if (XmlEl.Attributes["w:before"] != null)
                    return float.Parse(XmlEl.Attributes["w:before"].Value) / 20;
                return 0;
            }
            set
            {
                if (value != -1)
                    XmlEl.SetAttribute("before", XmlEl.NamespaceURI, ((int)(value * 20)).ToString());
            }
        }

        /// <summary>
        /// Отступ после абзаца
        /// Значение в пт (как в MS Word)
        /// </summary>
        public float After
        {
            get
            {
                if (XmlEl.Attributes["w:after"] != null)
                    return float.Parse(XmlEl.Attributes["w:after"].Value) / 20;
                return 0;
            }
            set
            {
                if (value != -1)
                    XmlEl.SetAttribute("after", XmlEl.NamespaceURI, ((int)(value * 20)).ToString());
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
            Ind.FirstLine = style.indentingFirtsLine;
            Ind.Hanging = style.indentingHanging;
            Ind.Left = style.indentingleft;
            Ind.Right = style.indentingRight;
            Spacing.Before = style.spacingBefore;
            Spacing.After = style.spacingAfter;
            Spacing.Line = style.spacingLine;
            PBorder.Left = style.borderLeft;
            PBorder.Right = style.borderRight;
            PBorder.Top = style.borderTop;
            PBorder.Bottom = style.borderBottom;
            PBorder.Between = style.borderBetween;
            PBorder.Bar= style.borderBar;
        }

        public bool HasSectPr
        {
            get
            {
                return ChildNodes.Where(x => x is SectProp).Any();
            }
        }

        public Ind Ind
        {
            get
            {
                return FindChildOrCreate<Ind>();
            }
        }
        /// <summary>
        /// Окончание секции. Следующая секция всегда начинается с новой страницы
        /// </summary>
        public SectProp SectPr
        {
            get
            {
                return FindChildOrCreate<SectProp>();
            }
        }

        /// <summary>
        /// Является элементом списка
        /// </summary>
        public bool HasNumPr
        {
            get
            {
                return FindChild<NumPr>()!=null;
            }
        }
        public NumPr NumPr
        {
            get
            {
                return FindChild<NumPr>();
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
            if (Ind.FirstLine != firtsLine || Ind.Hanging != hanging || Ind.Left != left || Ind.Right != right)
            {
                CreateChangeNode("w:pPrChange", XmlEl, author);
                Ind.FirstLine = firtsLine;
                Ind.Hanging = hanging;
                Ind.Left = left;
                Ind.Right = right;
            }
        }


        public void CompareBorder(Border left, Border right, Border top, Border bottom, Border between, Border bar, string author = "TDV")
        {
            PBorder.CompareBorder(BORDER.LEFT, left, author);
            PBorder.CompareBorder(BORDER.RIGHT, right, author);
            PBorder.CompareBorder(BORDER.TOP, top, author);
            PBorder.CompareBorder(BORDER.BOTTOM, bottom, author);
            PBorder.CompareBorder(BORDER.BETWEEN, between, author);
            PBorder.CompareBorder(BORDER.BAR, bar, author);
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
                CreateChangeNode("w:pPrChange", (XmlElement)XmlEl, author);
                NumPr.Level = level;
                NumPr.NumId.Value = (int)numId;
            } else if(!HasNumPr && numId!=null)
            {
                CreateChangeNode("w:pPrChange", (XmlElement)XmlEl, author);
                NewNodeLast<NumPr>();
                NumPr.Level = level;
                NumPr.NumId.Value = (int)numId;
            }
        }

        /// <summary>
        /// отступы
        /// </summary>
        public Spacing Spacing
        {
            get
            {
                return FindChildOrCreate<Spacing>();
            }
        }
        public void CompareSpacing(float before = -1, float after = -1, float line = -1, string author = "TDV")
        {
            if (Spacing.Before != before || Spacing.After != after || Spacing.Line != line)
            {
                CreateChangeNode("w:pPrChange", XmlEl, author);
                Spacing.Before = before;
                Spacing.After = after;
                Spacing.Line = line;
            }
        }


        public RProp RProp


        {
            
            
            
            
            
            
            
            
            
            
            
            
            get
            {
                return FindChildOrCreate<RProp>();
            }
        }

        public PBorder PBorder
        {
            get
            {
                return FindChildOrCreate<PBorder>();
            }
        }


        public void CompareHorizontalAlign(HORIZONTAL_ALIGN horizontalAlign, string author = "TDV")
        {
            if (HorizontalAlign != horizontalAlign)
            {
                CreateChangeNode("w:pPrChange", XmlEl, author);
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
                return FindChild<RProp>()?.IsBold ?? false;
            }
        }

        public bool IsItalic
        {
            get
            {
                return FindChild<RProp>()?.IsItalic ?? false;
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
                switch (XmlEl.GetAttribute("w:val"))
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
                        XmlEl.SetAttribute("val", XmlEl.NamespaceURI, "left");
                        break;
                    case HORIZONTAL_ALIGN.CENTER:
                        XmlEl.SetAttribute("val", XmlEl.NamespaceURI, "center");
                        break;
                    case HORIZONTAL_ALIGN.RIGHT:
                        XmlEl.SetAttribute("val", XmlEl.NamespaceURI, "right");
                        break;
                    case HORIZONTAL_ALIGN.BOTH:
                        XmlEl.SetAttribute("val", XmlEl.NamespaceURI, "both");
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

        public string RsidR
        {
            get
            {
                try
                {
                    return GetAttribute("w:rsidR");
                }
                catch (KeyNotFoundException)
                {
                    return null;
                }
            }
            set
            {
                if (string.IsNullOrEmpty(value))
                    RemoveAttribute("w:rsidR");
                else
                    SetAttribute("w:rsidR", value);
            }
        }

        public Size Width
        {
            get
            {
                Size result = new Size(0);
                FontStyle style = FontStyle.Regular;
                if (RProp.IsBold)
                    style = FontStyle.Bold;
                else if (RProp.IsItalic)
                    style = FontStyle.Italic;
                Font f = new Font(RProp.Font, (float)(RProp.FontSize??11),style);
                result.ValuePoints = (int)(TextRenderer.MeasureText(Text, f).Width*14.046d);
                return result;
            }
        }

        public RProp RProp
        {
            get
            {
                return FindChildOrCreate<RProp>();
            }
            set
            {
                RProp rProp = FindChild<RProp>();
                if (rProp != null)
                    XmlEl.RemoveChild(XmlEl);
                XmlEl.AppendChild(value.CopyXmlElement());
            }
        }

        /// <summary>
        /// Рецензирование - удаление
        /// создает тег w:del  и помещает в него эту ноду
        /// </summary>
        public void CorrectDel(string author = "TDV")
        {
            Del delNode = Parent.NewNodeBefore<Del>(XmlEl);
            delNode.Author = author;

            //delNode.XmlEl.AppendChild(CopyXmlElement());
            string text = Text;
            this.MoveTo(delNode);
            delNode.FindChild<R>().FindChildOrCreate<DelText>().Value = text;
            delNode.FindChild<R>().FindChild<T>()?.Delete();

            
            //Delete();
        }

        public void CorrectSetText(string newText, string author = "TDV")
        {
            if (Text != newText)
            {
                Ins newIns = Parent.NewNodeAfter<Ins>(XmlEl);
                newIns.Author = author;
                R newRnode = newIns.NewNodeLast<R>();
                newRnode.RProp = RProp;
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
                if (value.StartsWith(" ") || value.EndsWith(" ") || value.Contains("  "))
                    t.XmlSpace = XML_SPACE.PRESERVE;
                t.Text = value;
            }
        }

        public T t
        {
            get
            {
                return FindChildOrCreate<T>();
            }
        }


        public bool IsBold
        {
            get {return RProp.IsBold;}
        }

        public bool IsItalic
        {
            get { return RProp.IsItalic; }
        }
    }

    public enum XML_SPACE { NONE, DEFAULT,PRESERVE }


    public class T : Node
    {
        public T() : base("w:t") { }
        public T(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:t")
        { }

        /// <summary>
        /// preserve,default
        /// </summary>
        public XML_SPACE XmlSpace
        {
            get
            {
                if (!HasAttribute("xml:space"))
                    return XML_SPACE.NONE;
                switch (GetAttribute("xml:space"))
                {
                    case "preserve":
                        return XML_SPACE.PRESERVE;
                    case "default":
                        return XML_SPACE.DEFAULT;
                }
                throw new NotImplementedException();
            }
            set
            {
                switch (value)
                {
                    case XML_SPACE.DEFAULT:
                        SetAttribute("xml:space", "default");
                        return;
                    case XML_SPACE.PRESERVE:
                        SetAttribute("xml:space", "preserve");
                        return;
                    case XML_SPACE.NONE:
                        RemoveAttribute("xml:space");
                        break;
                }
                throw new NotImplementedException();
            }
        }

        public new string Text
        {
            get
            {
                return base.Text;
            }
            set
            {
                if (XmlEl != null)
                    XmlEl.InnerText = value;
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
    public class DelText : Node
    {
        public DelText() : base("w:delText") { }
        public DelText(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:delText")
        { }
        public string Value
        {
            get
            {
                return XmlEl.InnerText;
            }
            set
            {
                XmlEl.InnerText = value;
            }
        }
    }

    public class Paragraph : Node
    {
        public Paragraph() : base("w:p") { }
        public Paragraph(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:p")
        { }

        public string RsidR
        {
            get
            {
                try
                {
                    return GetAttribute("w:rsidR");
                }
                catch (KeyNotFoundException)
                {
                    return null;
                }
            }
            set
            {
                if (string.IsNullOrEmpty(value))
                    RemoveAttribute("w:rsidR");
                else
                    SetAttribute("w:rsidR", value);
            }
        }


        public string RsidRPr
        {
            get
            {
                try
                {
                    return GetAttribute("w:rsidRPr");
                }
                catch (KeyNotFoundException)
                {
                    return null;
                }
            }
            set
            {
                if (string.IsNullOrEmpty(value))
                    RemoveAttribute("w:rsidRPr");
                else
                    SetAttribute("w:rsidRPr", value);
            }
        }
        public string RsidRDefault
        {
            get
            {
                try
                {
                    return GetAttribute("w:rsidRDefault");
                }
                catch (KeyNotFoundException)
                {
                    return null;
                }
            }
            set
            {
                if (string.IsNullOrEmpty(value))
                    RemoveAttribute("w:rsidRDefault");
                else
                    SetAttribute("w:rsidRDefault", value);
            }
        }

        

        public Size Height
        {
            get
            {
                if (Section?.SectProp?.WorkspaceWidth == null)
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
            foreach(R r in RNodes)
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
            for(int rIndex=0;rIndex<RNodes.Count();rIndex++)
            {
                R r = RNodes[rIndex];
                Highlight currHighlight = r.FindChild<RProp>()?.FindChild<Highlight>();

                if(currHighlight != null)
                {
                    if (highlightText.Color == null)
                        highlightText.Color = currHighlight.Value;
                    //если подсветка следующей ноды такая же, как у текузей - зааппендить текст. иначе - доавбить в результат и создать новый HighlightText
                    if (rIndex + 1 <= RNodes.Count() - 1 && RNodes[rIndex+1].FindChild<RProp>()?.FindChild<Highlight>()?.Value==currHighlight.Value)
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
                foreach (Node insNode in ins.ChildNodes)
                    insNode.MoveTo(this);
                ins.Delete();
            }
            FindChild<PProp>()?.FindChild<PprChange>()?.Delete();
            FindChild<PProp>()?.FindChild<RProp>()?.FindChild<RprChange>()?.Delete();
            FindChild<PProp>()?.FindChild<RProp>()?.FindChild<Ins>()?.Delete();
            FindChild<PProp>()?.FindChild<SectProp>()?.FindChild<SectPrChange>()?.Delete();
            foreach (R r in RNodes)
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
                foreach (R r in RNodes)
                {
                    r.RProp.CompareStyle(style, author);
                }
            }
            PProp.RProp.CompareStyle(style, author);
        }

        public void ComparePStyle(PStyle style, string author = "TDV")
        {
            PProp.CompareStyle(style, author);
        }
        public void CompareStyles(PStyle pstyle,RStyle rstyle, string author = "TDV", bool applyToRnodes = true)
        {
            PProp.CompareStyle(pstyle, author);
            CompareRStyle(rstyle, author, applyToRnodes);
        }

        public List<R> RNodes
        {
            get { return FindChilds<R>(); }
        }

        public PProp PProp
        {
            get
            {
                return FindChildOrCreate<PProp>();
            }
        }

        public int DrawingCount()
        {
            int result = 0;
            foreach (R r in RNodes)
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
                return string.Join("", RNodes.Select(x => x.Text).ToList());
            }
            set
            {
                var rN = RNodes.FirstOrDefault();//тут может быть эксепшн, если нет ни одной r ноды. но это не точно
                if (rN == null)
                    rN=NewNodeLast<R>();
                foreach (var item in RNodes)
                    XmlEl.RemoveChild(item.XmlEl);
                rN.Text = value;
                XmlEl.AppendChild(rN.XmlEl);
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
                foreach (R r in RNodes)
                    r.CorrectDel();
                Ins newIns = NewNodeLast<Ins>();
                newIns.Author = author;
                R newRnode = newIns.NewNodeLast<R>();
                newRnode.RProp.SetStyle(rStyle);
                newRnode.Text = newText;
            }
        }

        /// <summary>
        /// Рецензирование - удаление
        /// создает тег w:del  и помещает в него эту ноду
        /// </summary>
        public void CorrectDel(string author = "TDV")
        {
            PProp.RProp.SetCorrectionMode("del");
            XmlEl.RemoveAttribute("w:rsidRPr");
            XmlEl.SetAttribute("rsidDel",XmlEl.NamespaceURI,XmlEl.GetAttribute("w:rsidR"));
            while (RNodes.Count() > 0)
                RNodes.First().CorrectDel(author);
            List<Paragraph> parList = Parent.FindChilds<Paragraph>();
            if (XmlEl == parList.Last().XmlEl)
            {
                //установить у ноды перед этой признак isDel. убирает лишний перенос на новую строку
                for (int i = parList.Count-1; i >= 0; i--)
                {
                        if (parList[i].XmlEl != XmlEl) //предыдущая нода
                        {
                            parList[i].PProp.RProp.SetCorrectionMode("del", author);
                        }
                        break;
                }
                Paragraph newP =  Parent.NewNodeAfter<Paragraph>(parList.Last().XmlEl);
                newP.XmlEl.SetAttribute("rsidRPr", XmlEl.NamespaceURI, XmlEl.GetAttribute("w:rsidR"));
            }
        }

        public bool AllRunsIsBold
        {
            get
            {
                bool result = true;
                if (RNodes.Count == 0)
                    return false;
                foreach (R run in RNodes)
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
                if (RNodes.Count == 0)
                    return false;
                foreach (R run in RNodes)
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
        public Size ParentSize;
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
            this.ParentSize = parentSize;
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
                        if (ParentSize == null)
                            throw new Exception("Для вычисления необходимо указать parentSize");
                        double percents = _value / 50 / 100;
                        return (int)(percents * ParentSize._value);
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

        public override int GetHashCode()
        {
            return _value.GetHashCode();
        }
        public override bool Equals(object obj)
        {
            if(obj is Size)
                return base.Equals((Size)obj);
            return false;
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

        public string RsidR
        {
            get
            {
                try
                {
                    return GetAttribute("w:rsidR");
                }
                catch (KeyNotFoundException)
                {
                    return null;
                }
            }
            set
            {
                if (string.IsNullOrEmpty(value))
                    RemoveAttribute("w:rsidR");
                else
                    SetAttribute("w:rsidR", value);
            }
        }


        public string RsidRPr
        {
            get
            {
                try
                {
                    return GetAttribute("w:rsidRPr");
                }
                catch (KeyNotFoundException)
                {
                    return null;
                }
            }
            set
            {
                if (string.IsNullOrEmpty(value))
                    RemoveAttribute("w:rsidRPr");
                else
                    SetAttribute("w:rsidRPr", value);
            }
        }

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
            if (FootnotePr.NumFmt.Value != fmt)
            {
                CreateChangeNode("w:sectPrChange",XmlEl, author);
                FootnotePr.NumFmt.Value = fmt;
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
                CreateChangeNode("w:sectPrChange", XmlEl, author);
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
                CreateChangeNode("w:sectPrChange", XmlEl, author);
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

            XmlElement header = (XmlElement)XmlEl.SelectSingleNode("w:headerReference[@w:type=\"" + stringType + "\"] ", Nsmgr);
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
                return docx.GetHeader(id);
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
                Override ov = docx.ContentTypes.GetOverride(newHeaderFile.GetFullPath(), true);
                ov.ContentType = Override.ContentTypes.HEADER;
                Relationship newRel = docx.WordRels.NewRelationship(newHeaderFile.Name, RELATIONSIP_TYPE.HEADER);
                Header newHeader = new Header(docx, newHeaderFile, newRel, create: true);
                //прописать в document.xml
                HeaderReference headerReference = docx.Document.Body.SectProp.GetHeaderReference(type, true);
                headerReference.Id = newRel.Id;
                docx.headers.Add(newRel.Id, newHeader);
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


            XmlElement footer = (XmlElement)XmlEl.SelectSingleNode("w:footerReference[@w:type=\"" + stringType + "\"] ", Nsmgr);
            DocxDocument docx = GetDocxDocument();
            if (footer != null)
            {
                string id = footer.GetAttribute("r:id");
                return docx.GetFooter(id);
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
                Override ov= docx.ContentTypes.GetOverride(newFooterFile.GetFullPath(), true);
                ov.ContentType = Override.ContentTypes.FOOTER;
                Relationship newRel= docx.WordRels.NewRelationship(newFooterFile.Name, RELATIONSIP_TYPE.FOOTER);
                Footer newFooter = new Footer(docx,newFooterFile,newRel, create:true);
                docx.footers.Add(newRel.Id, newFooter);
                //прописать в document.xml
                FooterReference footerReference= docx.Document.Body.SectProp.GetFooterReference(type, true);
                footerReference.Id = newRel.Id;
                //newFooter.Apply();
                return newFooter;
            }
        }

        public bool IsTitlePg
        {
            get
            {
                bool result = ChildNodes.Where(x => x.XmlEl.Name == "titlePg").Any();
                return result;
            }
            set
            {
                if (value)
                    if (!ChildNodes.Where(x => x.XmlEl.Name == "titlePg").Any())
                    {
                        XmlEl.AppendChild(XmlDoc.CreateElement("w:titlePg", XmlDoc.DocumentElement.NamespaceURI));
                    }
                    else
                    {
                        XmlElement forDel = ChildNodes.Where(x => x.XmlEl.Name == "titlePg").FirstOrDefault()?.XmlEl;
                        if (forDel != null)
                            XmlDoc.RemoveChild(forDel);
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
                CreateChangeNode("w:pPrChange", (XmlElement)XmlEl.ParentNode, author);
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
            XmlElement n = (XmlElement)XmlEl.SelectSingleNode($"{prefix}:{localName}", Nsmgr);
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
            XmlElement n = (XmlElement)XmlEl.SelectSingleNode($"{prefix}:{localName}", Nsmgr);
            if (b.type == LINE_TYPE.NONE && n != null)
            {
                XmlEl.RemoveChild(n);
            }

            if (b.type != LINE_TYPE.NONE)
            {
                if (n == null)
                {
                    n = XmlDoc.CreateElement($"{prefix}:{localName}", XmlDoc.DocumentElement.NamespaceURI);
                    XmlEl.AppendChild(n);
                }
                switch (b.type)
                {
                    case LINE_TYPE.SINGLE:
                        n.SetAttribute("val", XmlEl.NamespaceURI, "single");
                        break;
                    default:
                        break;
                }
                n.SetAttribute("sz", XmlEl.NamespaceURI, b.size.ToString());
                n.SetAttribute("space", XmlEl.NamespaceURI, b.space.ToString());
                n.SetAttribute("color", XmlEl.NamespaceURI, b.color);
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
            if (Border != b)
            {
                CreateChangeNode("w:rPrChange", (XmlElement)XmlEl.ParentNode, author);
                Border = b ;
            }
        }

        public Border Border
        {
            get
            {
                Border b = new Border();
                if(XmlEl.HasAttribute("w:color"))
                    b.color = XmlEl.GetAttribute("w:color");
                if(XmlEl.HasAttribute("w:sz"))
                    b.size = Int32.Parse(XmlEl.GetAttribute("w:sz"));
                if (XmlEl.HasAttribute("w:space"))
                    b.space = Int32.Parse(XmlEl.GetAttribute("w:space"));
                else b.space = -1;
                switch (XmlEl.GetAttribute("w:val"))
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
                    Parent.XmlEl.RemoveChild(XmlEl);
                    return;
                }

                else
                {
                    switch ( value.type)
                    {
                        case LINE_TYPE.SINGLE:
                            XmlEl.SetAttribute("val", XmlEl.NamespaceURI, "single");
                            break;
                        default:
                            break;
                    }
                    XmlEl.SetAttribute("sz", XmlEl.NamespaceURI, value.size.ToString());
                    XmlEl.SetAttribute("space", XmlEl.NamespaceURI, value.space.ToString());
                    XmlEl.SetAttribute("color", XmlEl.NamespaceURI, value.color);
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

        public int Id
        {
            get
            {
                return Int32.Parse(GetAttribute("w:id"));
            }
            set
            {
                SetAttribute("w:id", value.ToString());
            }
        }

        public DateTime? Date
        {
            get
            {
                try
                {
                    return DateTime.Parse(GetAttribute("w:date"));
                }
                catch
                {
                    return null;
                }
            }
            set
            {
                if (value == null)
                    RemoveAttribute("w:date");
                else
                    SetAttribute("w:date", ((DateTime)value).ToString("yyyy-MM-ddTHH:mm:ssZ"));
            }
        }

        /// <summary>
        /// Автор комментария
        /// </summary>
        public string Author
        {
            get
            {
                return GetAttribute("w:author");
            }
            set
            {
                SetAttribute("w:author", value);
            }
        }

        //метод общий для INS и DEL
        public override void InitXmlElement()
        {
            base.InitXmlElement();
            Id = GetDocxDocument().Document.GetLastId() + 1;
            Author = "TDV";
            Date = DateTime.Now;
        }
    }

    /// <summary>
    /// Размер шрифта в половниах точек или 1/144 дюйма
    /// </summary>
    public class Sz : Node
    {
        public Sz() : base("w:sz") { }

        public Sz(Node parent) : base(parent, "w:sz") { }

        public Sz(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:sz") { }
        public double? Value
        {
            get
            {
                try
                {
                    return double.Parse(GetAttribute("w:val"));
                }
                catch { return null; }
            }
            set
            {
                if (value == null)
                    Delete();
                else
                    SetAttribute("w:val", value.ToString());
            }
        }
    }

    /// <summary>
    /// размер шрифта для сложных наборов знаков. размер в полуточках или 1/144 дюйма
    /// </summary>
    public class SzCs : Node
    {
        public SzCs() : base("w:szCs") { }

        public SzCs(Node parent) : base(parent, "w:szCs") { }

        public SzCs(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:szCs") { }
        public double? Value
        {
            get
            {
                try
                {
                    return double.Parse(GetAttribute("w:val"));
                }
                catch { return null; }
            }
            set
            {
                if (value == null)
                    Delete();
                else
                    SetAttribute("w:val", value.ToString());
            }
        }
    }

    public class Del : Node
    {
        public Del() : base("w:del") { }

        public Del(Node parent) : base(parent, "w:del") { }

        public Del(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:del") { }

        public int Id
        {
            get
            {
                return Int32.Parse(GetAttribute("w:id"));
            }
            set
            {
                SetAttribute("w:id", value.ToString());
            }
        }

        public DateTime? Date
        {
            get
            {
                try
                {
                    return DateTime.Parse(GetAttribute("w:date"));
                }
                catch
                {
                    return null;
                }
            }
            set
            {
                if (value == null)
                    RemoveAttribute("w:date");
                else
                    SetAttribute("w:date", ((DateTime)value).ToString("yyyy-MM-ddTHH:mm:ssZ"));
            }
        }

        /// <summary>
        /// Автор
        /// </summary>
        public string Author
        {
            get
            {
                return GetAttribute("w:author");
            }
            set
            {
                SetAttribute("w:author", value);
            }
        }

        //метод общий для INS и DEL
        public override void InitXmlElement()
        {
            base.InitXmlElement();
            Id = GetDocxDocument().Document.GetLastId() + 1;
            Author = "TDV";
            Date = DateTime.Now;
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
                        .FindChild<Extent>().Cx/360000.0f;
                }
                catch
                { return -1; }
            }
            set
            {
                try
                {
                    FindChild<Inline>()
                        .FindChild<Extent>().Cx = (int)(value * 360000.0f);
                    FindChild<Inline>().FindChild<EffectExtent>().L = 0;
                    FindChild<Inline>().FindChild<EffectExtent>().R = 0;
                    FindChild<Inline>().FindChild<EffectExtent>().B = 0;
                    FindChild<Inline>().FindChild<EffectExtent>().T = 0;
                    FindChild<Inline>()
                        .FindChild<Graphic>().FindChild<GraphicData>().FindChild<Pic>().FindChild<SpPr>().FindChild<Xfrm>().FindChild<Ext>().Cx = (int)(value * 360000.0f);
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
                        .FindChild<Extent>().Cy/ 360000.0f;
                }
                catch
                { return -1; }
            }
            set
            {
                try
                {
                    FindChild<Inline>()
                        .FindChild<Extent>().Cy = (int)(value * 360000.0f);
                    FindChild<Inline>().FindChild<EffectExtent>().L = 0;
                    FindChild<Inline>().FindChild<EffectExtent>().R = 0;
                    FindChild<Inline>().FindChild<EffectExtent>().B = 0;
                    FindChild<Inline>().FindChild<EffectExtent>().T = 0;
                    FindChild<Inline>()
                        .FindChild<Graphic>().FindChild<GraphicData>().FindChild<Pic>().FindChild<SpPr>().FindChild<Xfrm>().FindChild<Ext>().Cy = (int)(value * 360000.0f);
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
        public int Cx
        {
            get
            {
                int result = -1;
                Int32.TryParse(XmlEl.GetAttribute("cx"), out result);
                return result;
            }
            set { XmlEl.SetAttribute("cx", value.ToString()); }
        }
        /// <summary>
        /// Высота
        /// </summary>
        public int Cy
        {
            get
            {
                int result = -1;
                Int32.TryParse(XmlEl.GetAttribute("cy"), out result);
                return result;
            }
            set{ XmlEl.SetAttribute("cy", value.ToString()); }
        }
    }

    public class EffectExtent : Node
    {
        public EffectExtent() : base("wp:effectExtent") { }
        public EffectExtent(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "wp:effectExtent") { }

        public int L
        {
            get
            {
                int result = -1;
                Int32.TryParse(XmlEl.GetAttribute("l"), out result);
                return result;
            }
            set { XmlEl.SetAttribute("l", value.ToString()); }
        }
        public int R
        {
            get
            {
                int result = -1;
                Int32.TryParse(XmlEl.GetAttribute("r"), out result);
                return result;
            }
            set { XmlEl.SetAttribute("r", value.ToString()); }
        }
        public int T
        {
            get
            {
                int result = -1;
                Int32.TryParse(XmlEl.GetAttribute("t"), out result);
                return result;
            }
            set { XmlEl.SetAttribute("t", value.ToString()); }
        }
        public int B
        {
            get
            {
                int result = -1;
                Int32.TryParse(XmlEl.GetAttribute("b"), out result);
                return result;
            }
            set { XmlEl.SetAttribute("b", value.ToString()); }
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
        public int Cx
        {
            get
            {
                int result = -1;
                Int32.TryParse(XmlEl.GetAttribute("cx"), out result);
                return result;
            }
            set { XmlEl.SetAttribute("cx", value.ToString()); }
        }
        /// <summary>
        /// Высота
        /// </summary>
        public int Cy
        {
            get
            {
                int result = -1;
                Int32.TryParse(XmlEl.GetAttribute("cy"), out result);
                return result;
            }
            set { XmlEl.SetAttribute("cy", value.ToString()); }
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
        public string Embed
        {
            get { return XmlEl.GetAttribute("r:embed"); }
        }

        public Image GetImageFile()
        {
            byte[] bytes = GetDocxDocument().WordRels.GetFileById(Embed).Content;
            return Bitmap.FromStream(new MemoryStream(bytes));
            
        }
    }

    public class FootnotePr : Node
    {
        public FootnotePr() : base("w:footnotePr") { }
        public FootnotePr(Node parent) : base(parent, "w:footnotePr") { }
        public FootnotePr(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:footnotePr") { }
        public NumFmt NumFmt{
            get
            {
                return FindChildOrCreate<NumFmt>();
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
                switch (XmlEl.GetAttribute("w:val"))
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
                        XmlEl.RemoveAttribute("val", XmlEl.NamespaceURI);
                        break;
                    case NUM_FMT.CHICAGO:
                        XmlEl.SetAttribute("val", XmlEl.NamespaceURI, "chicago");
                        break;
                    case NUM_FMT.LOWER_ROMAN:
                        XmlEl.SetAttribute("val", XmlEl.NamespaceURI, "lowerRoman");
                        break;
                    case NUM_FMT.UPPER_ROMAN:
                        XmlEl.SetAttribute("val", XmlEl.NamespaceURI, "upperRoman");
                        break;
                    case NUM_FMT.UPPER_LETTER:
                        XmlEl.SetAttribute("val", XmlEl.NamespaceURI, "upperLetter");
                        break;
                    case NUM_FMT.BULLET:
                        XmlEl.SetAttribute("val", XmlEl.NamespaceURI, "bullet");
                        break;
                }
            }
        }
    }



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
                return Int32.Parse(XmlEl.GetAttribute("w:val"));
            }
            set
            {
                XmlEl.SetAttribute("val", XmlEl.NamespaceURI, value.ToString());
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
                return  Int32.Parse(XmlEl.GetAttribute("w:val"));
            }
            set
            {
                XmlEl.SetAttribute("val", XmlEl.NamespaceURI, value.ToString());
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
                Ilvl result = ChildNodes.Where(x => x is Ilvl).Select(x => (Ilvl)x).FirstOrDefault();
                if (result == null)
                    result = new Ilvl(this);
                return result;
            }
        }
        public NumId NumId
        {
            get
            {
                NumId result = ChildNodes.Where(x => x is NumId).Select(x => (NumId)x).FirstOrDefault();
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

    public class Hyperlink : Node
    {
        public Hyperlink() : base("w:hyperlink") { }
        public Hyperlink(Node parent, int numId) : base(parent, "w:hyperlink")   {}
        public Hyperlink(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:hyperlink") { }

        public List<R> RNodes
        {
            get
            {
                return FindChilds<R>();
            }
        }

        public override string Text { 
            get => base.Text;
            set
            {
                while (RNodes.Count > 0)
                    RNodes.First().Delete();
                DocxDocument docx = GetDocxDocument();
                Style hlStyle = docx.Styles.GetStyleByName("Hyperlink");
                if(hlStyle==null)
                {
                    hlStyle = docx.Styles.NewNodeLast<Style>();
                    hlStyle.Type = STYLE_TYPE.CHARACTER;
                    hlStyle.StyleId = $"a{ docx.Styles.GetMaxStyleId("a") + 1}";
                    hlStyle.Name = "Hyperlink";
                    hlStyle.basedOn = docx.Styles.GetDefaultParagraphFontStyle();
                    hlStyle.UiPriority = 99;
                    hlStyle.IsUnhideWhenUsed = true;
                    RProp rPr = hlStyle.FindChildOrCreate<RProp>();
                    rPr.Color = "0563C1";
                    rPr.Underline = LINE_TYPE.SINGLE;
                }
                R r = NewNodeLast<R>();
                r.RProp.Style = hlStyle;
                r.t.XmlSpace= XML_SPACE.PRESERVE;
                r.Text = value;
            }
        }

        /// <summary>
        /// Ссылка на Relationships.xml
        /// </summary>
        public string Id
        {
            get
            {
                if(HasAttribute("r:id"))
                    try { 
                        return GetAttribute("r:id");
                    }catch{}

                return null;
            }
            set
            {
                SetAttribute("r:id", value);
            }
        }

        /// <summary>
        /// ссылка на закладку в документе
        /// </summary>
        public string Anchor
        {
            get
            {
                try
                {
                    return GetAttribute("r:anchor");
                }
                catch (KeyNotFoundException)
                {
                    return null;
                }
            }
            set
            {
                SetAttribute("r:anchor", value);
            }
        }

        public int? History
        {
            get
            {
                try
                {
                    return Int32.Parse(GetAttribute("r:id"));
                }
                catch
                {
                    return null;
                }
            }
            set
            {
                if (value == null)
                    RemoveAttribute("r:history");
                else
                    SetAttribute("r:history", ((int)value).ToString());
            }
        }

        public Relationship GetRelationship(bool createIfNotExist=false,RELATIONSIP_TYPE newRelType=RELATIONSIP_TYPE.HYPERLINK)
        {
            DocxDocument docx = GetDocxDocument();
            if (string.IsNullOrEmpty(Id))
            {
                if (!createIfNotExist)
                    throw new KeyNotFoundException($"Id не заполнен. Невозможно вычислить связь");
                Id = $"rId{docx.WordRels.GetMaxRelId() + 1}";
                Relationship rel = docx.WordRels.NewRelationship("", newRelType);
                //docx.wordRels.Apply();
                return rel;
            }
            else
            {
                return docx.WordRels.GetRelationshipById(Id);
            }
        }

        public string Url
        {
            get
            {
                return GetRelationship().Target;
            }
            set
            {
                GetRelationship(true).Target = value;
                GetDocxDocument().WordRels.Apply();
            }
        }
    }
    /// <summary>
    /// подчеркивание (underline)
    /// </summary>
    public class U : Node
    {
        public U() : base("w:u") { }
        public U(Node parent, int numId) : base(parent, "w:u") { }
        public LINE_TYPE Value
        {
            get
            {
                if (!HasAttribute("w:val"))
                    return LINE_TYPE.NONE;
                switch (GetAttribute("w:val"))
                {
                    case "dotted":
                        return LINE_TYPE.DOTTED;
                    case "single":
                        return LINE_TYPE.SINGLE;
                    default:
                        throw new NotImplementedException();
                }
            }
            set
            {
                switch (value)
                {
                    case LINE_TYPE.NONE:
                        Delete();
                        return;
                    case LINE_TYPE.SINGLE:
                        SetAttribute("w:val", "single");
                        return;
                    case LINE_TYPE.DOTTED:
                        SetAttribute("w:val", "dotted");
                        return;
                }
                throw new NotImplementedException();

            }
        }
    }
    public class Lang : Node
    {
        public Lang() : base("w:lang") {     }
        public Lang(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:lang") { }

        public override void InitXmlElement()
        {
            base.InitXmlElement();
            Value = "en-US";
        }
        public string Value
        {
            get
            {
                try
                {
                    return GetAttribute("w:val");
                }
                catch (KeyNotFoundException)
                {
                    return null;
                }
            }
            set
            {
                if (string.IsNullOrEmpty(value))
                    RemoveAttribute("w:val");
                else
                    SetAttribute("w:val", value);
            }
        }
    }
}
