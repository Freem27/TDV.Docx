using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml;
using System.Linq;
using System.Runtime.CompilerServices;
[assembly:InternalsVisibleTo("TDVDocx.Test")]

namespace TDV.Docx
{
    public class PStyle
    {
        public PStyle(HORIZONTAL_ALIGN horizontalAlign,
            Border borderLeft,
            Border borderRight,
            Border borderTop,
            Border borderBottom,
            Border borderBetween,
            Border borderBar,
            float spacingBefore,
            float spacingAfter,
            float spacingLine,
            float indentingFirtsLine,
            float indentingHanging,
            float indentingleft, float indentingRight)
        {
            this.horizontalAlign = horizontalAlign;
            this.borderLeft = borderLeft;
            this.borderRight = borderRight;
            this.borderTop = borderTop;
            this.borderBottom = borderBottom;
            this.borderBetween = borderBetween;
            this.borderBar = borderBar;
            this.spacingBefore = spacingBefore;
            this.spacingAfter = spacingAfter;
            this.spacingLine = spacingLine;
            this.indentingFirtsLine = indentingFirtsLine;
            this.indentingHanging = indentingHanging;
            this.indentingleft = indentingleft;
            this.indentingRight = indentingRight;
        }
        public HORIZONTAL_ALIGN horizontalAlign;
        public Border borderLeft;
        public Border borderRight;
        public Border borderTop;
        public Border borderBottom;
        public Border borderBetween;
        public Border borderBar;
        public float spacingBefore;
        public float spacingAfter;
        public float spacingLine;
        public float indentingFirtsLine;
        public float indentingHanging;
        public float indentingleft;
        public float indentingRight;

        public PStyle Copy()
        {
            return new PStyle(horizontalAlign, borderLeft, borderRight, borderTop, borderBottom,borderBetween,borderBar, spacingBefore,
                spacingAfter, spacingLine, indentingFirtsLine, indentingHanging, indentingleft, indentingRight);
        }
    }

    public class RStyle
    {
        public RStyle(bool isBold,
            string font,
            float fontSize,
            bool isItalic,
            bool isStrike,
            LINE_TYPE underline,
            string color,
            string highlight,
            Border border)
        {
            this.isBold = isBold;
            this.font = font;
            this.fontSize = fontSize;
            this.isItalic = isItalic;
            this.isStrike = isStrike;
            this.underline = underline;
            this.color = color;
            this.highlight = highlight;
            this.border = border;
        }
        public bool isBold;
        public string font;
        public float fontSize;
        public bool isItalic;
        public bool isStrike;
        public LINE_TYPE underline;
        public string color;
        public string highlight;
        public Border border;
        public RStyle Copy()
        {
            return new RStyle(isBold,font,fontSize,isItalic,isStrike,underline,color,highlight,border);
        }
    }

    internal static class DocXExtentions
    {
       public static int GetLastId(this XmlDocument doc)
       {
           int result = -1; 
           XmlNamespaceManager nsmgr = new XmlNamespaceManager(doc.NameTable);
            nsmgr.AddNamespace("w", doc.DocumentElement.NamespaceURI);
            XmlNodeList insDelList = doc.SelectNodes("//*[@w:id]", nsmgr);
           if (insDelList.Count>0)
               foreach (XmlNode item in insDelList)
               {
                   XmlElement el = (XmlElement)item;
                   int elId = int.Parse(el.GetAttribute("id", el.NamespaceURI));
                   if (result < elId)
                   {
                       result = elId;
                   }
               }
            return result;
        }
    }
    public enum HORIZONTAL_ALIGN {NONE,LEFT,CENTER,RIGHT,BOTH}

    public enum BORDER
    {
        LEFT,RIGHT,TOP,BOTTOM,
        /// <summary>
        /// Для параграфов. Между параграфами
        /// </summary>
        BETWEEN,
        BAR,
        /// <summary>
        /// Для таблиц. Все внутренние горизонтальные
        /// </summary>
        INSIDE_H,
        /// <summary>
        /// Для таблиц. Все внутренние вертикальные
        /// </summary>
        INSIDE_V
    }
    public enum LINE_TYPE {NONE, SINGLE,DOTTED,UNKNOWN }

    /// <summary>
    /// тип связи. для заголовков\футеров
    /// </summary>
    public enum REFERENCE_TYPE { FIRST,EVEN,DEFAULT}

    public class Node
    {
        protected Node(string qualifiedName="")
        {
            this.qualifiedName = qualifiedName;
        }

        public XmlElement CopyXmlElement()
        {
            XmlElement result = doc.CreateElement(xmlEl.Name,xmlEl.NamespaceURI);
            result.InnerXml = xmlEl.InnerXml;
            foreach (XmlAttribute att in xmlEl.Attributes)
                result.SetAttribute(att.LocalName,xmlEl.NamespaceURI, att.Value);
            return result;
        }

        public Node(Node parent,string qualifiedName="")
        {
            this.parent = parent;
            this.qualifiedName = qualifiedName;
            this.nsmgr = parent.nsmgr;
            this.doc = parent.doc;
            InitXmlElement();
            if (this is PProp || this is RProp)
                parent.xmlEl.InsertBefore(xmlEl,parent.xmlEl.FirstChild);
            else
                parent.xmlEl.AppendChild(xmlEl);
        }
        public Node(XmlElement xmlElement, Node parent, string qualifiedName) : this(qualifiedName)
        {
            this.parent = parent;
            doc = xmlElement.OwnerDocument;
            nsmgr = parent.nsmgr;
            xmlEl = xmlElement;
        }
        public XmlDocument doc;
        public XmlNamespaceManager nsmgr;
        public Node parent;

        protected string qualifiedName;

        /// <summary>
        /// 
        /// </summary>
        internal List<Node> baseStyleNodes = new List<Node>();

        public T GetParentRecurcieve<T>() where T:Node
        {
            if (parent == null)
                return null;
            if (parent is T)
                return (T) parent;
            else
                return parent.GetParentRecurcieve<T>();
        }

        public T FindChild<T>() where T : Node
        {
            return (T)childNodes.Where(x => x is T).FirstOrDefault();
        }

        public List<T> FindChilds<T>() where T : Node
        {
            return childNodes.Where(x => x is T).Select(x=>(T)x).ToList();
        }

        public List<T> FindChildsRecurcieve<T>() where T : Node
        {
            List<T> result = new List<T>();
            result= childNodes.Where(x => x is T).Select(x => (T)x).ToList();
            foreach (Node child in childNodes)
            {
                result.AddRange(child.FindChildsRecurcieve<T>());
            }
            return result;
        }

        public virtual List<Node> childNodes
        {
            get
            {
                List<Node> result = new List<Node>();
                if (xmlEl == null)
                    return result;
                foreach (var el in xmlEl.ChildNodes)
                {
                    if (!(el is XmlElement))
                        continue;
                    XmlElement item = (XmlElement)el;
                    switch (item.Name)
                    {
                        case "w:body":
                            result.Add(new Body(item, this));
                            break;
                        case "w:style":
                            result.Add(new Style(item, this));
                            break;
                        case "w:rPr":
                            result.Add(new RProp(item, this));
                            break;
                        case "w:bdr":
                            result.Add(new RBorder(item, this));
                            break;
                        case "w:r":
                            result.Add(new R(item, this));
                            break;
                        case "w:p":
                            result.Add(new Paragraph(item, this));
                            break;
                        case "w:pPr":
                            result.Add(new PProp(item, this));
                            break;
                        case "w:sectPr":
                            result.Add(new SectProp(item, this));
                            break;
                        case "w:del":
                            result.Add(new Del(item, this));
                            break;
                        case "w:tbl":
                            result.Add(new Table(item, this));
                            break;
                        case "w:tblPr":
                            result.Add(new TableProp(item, this));
                            break;
                        case "w:tblStyle":
                            result.Add(new TblStyle(item, this));
                            break;
                        case "w:tblBorders":
                            result.Add(new TblBorders(item, this));
                            break;
                        case "w:tblW":
                            result.Add(new TableWidth(item, this));
                            break;
                        case "w:tblInd":
                            result.Add(new TblInd(item, this));
                            break;
                        case "w:tblGrid":
                            result.Add(new TableGrid(item, this));
                            break;
                        case "w:gridCol":
                            result.Add(new GridColumn(item, this, result.Where(x => x is GridColumn).Count()));
                            break;
                        case "w:tr":
                            int trInd = result.Where(x => x is Tr).Count();
                            Tr newTr = new Tr(item, this, trInd);
                            result.Add(newTr);
                            //result.Add(new Tr(item, this));
                            break;
                        case "w:trPr":
                            result.Add(new TrProp(item, this));
                            break;
                        case "w:trHeight":
                            result.Add(new TrHeight(item, this));
                            break;
                        case "w:tc":
                            int colInd = result.Where(x => x is Tc).Count();
                            Tc newTc = new Tc(item, this,colInd);
                            result.Add(newTc);
                            break;
                        case "w:tcPr":
                            result.Add(new TcProp(item, this));
                            break;
                        case "w:vAlign":
                            result.Add(new VAlign(item, this));
                            break;
                        case "w:tcBorders":
                            result.Add(new TcBorders(item, this));
                            break;
                        case "w:tcW":
                            result.Add(new TcW(item, this));
                            break;
                        case "w:vMerge":
                            result.Add(new VMerge(item, this));
                            break;
                        case "w:gridSpan":
                            result.Add(new GridSpan(item, this));
                            break;
                        case "w:pBdr":
                            result.Add(new PBorder(item, this));
                            break;
                        case "w:ind":
                            result.Add(new Ind(item, this));
                            break;
                        case "w:pgMar":
                            result.Add(new PageMargin(item, this));
                            break;
                        case "w:spacing":
                            result.Add(new Spacing(item, this));
                            break;
                        
                        case "w:drawing":
                            result.Add(new Drawing(item, this));
                            break;
                        case "wp:inline":
                            result.Add(new Inline(item, this));
                            break;
                        case "wp:extent":
                            result.Add(new Extent(item, this));
                            break;
                        case "wp:effectExtent":
                            result.Add(new EffectExtent(item, this));
                            break;
                        case "a:xfrm":
                            result.Add(new Xfrm(item, this));
                            break;
                        case "a:ext":
                            result.Add(new Ext(item, this));
                            break;
                        case "a:graphic":
                            result.Add(new Graphic(item, this));
                            break;
                        case "a:graphicData":
                            result.Add(new GraphicData(item, this));
                            break;
                        case "pic:pic":
                            result.Add(new Pic(item, this));
                            break;
                        case "pic:spPr":
                            result.Add(new SpPr(item, this));
                            break;
                        case "pic:blipFill":
                            result.Add(new BlipFill(item, this));
                            break;
                        case "a:blip":
                            result.Add(new Blip(item, this));
                            break;
                        case "w:footnote":
                            result.Add(new Footnote(item, this));
                            break;
                        case "w:footnotePr":
                            result.Add(new FootnotePr(item, this));
                            break;
                        case "w:numFmt":
                            result.Add(new NumFmt(item, this));
                            break;
                        default:
                            result.Add(new Node(item, this, item.Name));
                            break;
                    }

                }

                result.AddRange(baseStyleNodes);
                return result;
            }
        }

        public void CreateChangeNode(string changeNodeName = "w:pPrChange", XmlElement moveChangeNodeTo = null, string author = "TDV")
        {
            XmlElement oldNode = this.CopyXmlElement();
            if (moveChangeNodeTo == null)
                moveChangeNodeTo = xmlEl;
            XmlElement nChange = (XmlElement)moveChangeNodeTo.SelectSingleNode(changeNodeName, nsmgr);
            //создать ноду w: rPrChange если она не создана
            if (nChange == null)
            {
                nChange = doc.CreateElement(changeNodeName, doc.DocumentElement.NamespaceURI);
                nChange.SetAttribute("id", xmlEl.NamespaceURI, (doc.GetLastId() + 1).ToString());
                moveChangeNodeTo.AppendChild(nChange);
            }
            if (nChange.SelectSingleNode(oldNode.Name, nsmgr) == null)
                nChange.AppendChild(oldNode); //Скопировать в нее этот rPr
            nChange.SetAttribute("author", xmlEl.NamespaceURI, author);
            nChange.SetAttribute("date", xmlEl.NamespaceURI, DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ssZ"));

            if (changeNodeName == "w:pPrChange")
            {
                var rprForDel = nChange.SelectSingleNode("w:rPr", nsmgr);
                if(rprForDel!=null)
                    nChange.RemoveChild(rprForDel);
            }else if (changeNodeName == "w:sectPrChange")
            {
                var sectChangeNode = nChange.SelectSingleNode("w:sectPr", nsmgr);
                foreach (XmlElement forDel in sectChangeNode.SelectNodes("w:headerReference", nsmgr))
                    sectChangeNode.RemoveChild(forDel);
                foreach (XmlElement forDel in sectChangeNode.SelectNodes("w:footerReference", nsmgr))
                    sectChangeNode.RemoveChild(forDel);

            }
        }

        public XmlElement xmlEl;
    
        public void Delete()
        {
            if (parent != null && parent.xmlEl.SelectSingleNode(xmlEl.Name, nsmgr)!=null)
            {
                parent.xmlEl.RemoveChild(xmlEl);
            }
        }

        public void MoveTo(Node nodeTo)
        {
            Delete();
            if(nodeTo.parent!=this)
                nodeTo.xmlEl.AppendChild(xmlEl);
        }
        
        public void Clear()
        {
            xmlEl.RemoveAll();
            childNodes.Clear();
        }

        //добавляет новую НОДУ в конец списка
        private T NewNode<T>() where T: Node
        {
            //Node result = new Node(xmlEl.OwnerDocument.CreateElement(""),this);
            T result = Activator.CreateInstance<T>();
            result.doc = xmlEl.OwnerDocument;
            result.parent = this;
            result.nsmgr = nsmgr;
            result.InitXmlElement();
            //xmlEl.AppendChild(result.xmlEl);
            return result;
        }

        public T NewNodeAfter<T>(XmlElement after) where T : Node
        {
            T result = NewNode<T>();
            xmlEl.InsertAfter(result.xmlEl, after);
            return result;
        }

        public T NewNodeBefore<T>(XmlElement before) where T : Node
        {
            T result = NewNode<T>();
            xmlEl.InsertBefore(result.xmlEl, before);
            return result;
        }

        public T NewNodeFirst<T>() where T : Node
        {
            T result = NewNode<T>();
            xmlEl.InsertBefore(result.xmlEl, xmlEl.FirstChild);
            return result;
        }

        public T NewNodeLast<T>() where T : Node
        {
            T result = NewNode<T>();
            xmlEl.AppendChild(result.xmlEl);
            return result;
        }


        /// <summary>
        /// Создает новый XmlElement. Необходимо переопределеять в классах наследниках
        /// </summary>
        public virtual void InitXmlElement()
        {
            xmlEl = doc.CreateElement(qualifiedName,doc.DocumentElement.NamespaceURI);
        }

        //public void Append(Node node)
        //{
        //    childNodes.Add(node);
        //    xmlEl.AppendChild(node.xmlEl);
        //}

        public virtual string Text
        {
            get
            {
                if(xmlEl!=null)
                    return xmlEl.InnerText;
                return null;
            }
        }

        public DocxDocument GetDocxDocument()
        {
            DocxDocument result = null;

            if (this is BaseNode)
                result = ((BaseNode)this).docxDocument;
            else if (parent != null)
                result = parent.GetDocxDocument();
            return result;
        }
    }

    public interface ICorrectable 
    {
        /// <summary>
        /// устанавливает режим правки
        /// mode = "del", "ins"
        /// </summary>
        void SetCorrectionMode(string mode, string author = "TDV");
    }
}
