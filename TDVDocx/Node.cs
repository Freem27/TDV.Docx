﻿using System;
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
            float indentingleft, float indentingRight,
            int? numId=null,
            int numLevel=0)
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
            this.numId = numId;
            this.numLevel = numLevel;
        }
        /// <summary>
        /// ссылка на формат списка
        /// </summary>
        public int numLevel;
        public int? numId;
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
       public static int GetLastId(this XmlDocument doc,int start=-1)
       {
           int result = 0; 
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

    public enum INSERT_POS { FIRST, LAST }
    public class Node
    {
        protected Node(string qualifiedName="")
        {
            this.qualifiedName = qualifiedName;
        }

        public Section Section
        {
            get
            {
                foreach (Section s in GetDocxDocument().document.Sections)
                    if (s.childNodes.Where(x => x.xmlEl == xmlEl).Any())
                        return s;
                return null;
            }
        }
        public Node NextNode;
        public Node PrevNode;
        public virtual void ApplyAllFixes()
        {
            foreach(Node n in childNodes)
            {
                n.ApplyAllFixes();
            }
        }

        public string GenerateGuid()
        {
            return Guid.NewGuid().ToString().Substring(0,8).ToUpper();
        }

        public int GenerateId(int len=9)
        {
            return new Random().Next((int)Math.Pow(10.0, (double)(len - 1)), (int)Math.Pow(10.0 , (double)len) - 1);
        }

        public XmlElement CopyXmlElement()
        {
            XmlElement result = xmlDoc.CreateElement(xmlEl.Name,xmlEl.NamespaceURI);
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
            this.xmlDoc = parent.xmlDoc;
            InitXmlElement();
            if (this is PProp || this is RProp)
                parent.xmlEl.InsertBefore(xmlEl,parent.xmlEl.FirstChild);
            else
                parent.xmlEl.AppendChild(xmlEl);
        }
        public Node(XmlElement xmlElement, Node parent, string qualifiedName) : this(qualifiedName)
        {
            this.parent = parent;
            xmlDoc = xmlElement.OwnerDocument;
            nsmgr = parent.nsmgr;
            xmlEl = xmlElement;
        }
        public XmlDocument xmlDoc;
        public XmlNamespaceManager nsmgr;
        public Node parent;

        protected string qualifiedName;


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

        internal string GetAttribute(string name)
        {
            string prefix = null;
            string localName = name;
            if(name.Contains(":"))
            {
                prefix = name.Split(':')[0];
                localName = name.Split(':')[1];
            }
            foreach(XmlAttribute a in xmlEl.Attributes)
            {
                if (a.Name == name || (a.LocalName==localName && a.Prefix==prefix))
                    return a.Value;
            }
            throw new KeyNotFoundException();
        }
        internal void SetAttribute(string name,string value)
        {
            string prefix = null;
            string localName = name;
            if (name.Contains(":"))
            {
                prefix = name.Split(':')[0];
                localName = name.Split(':')[1];
            }

            if (xmlEl.HasAttribute(localName))
                xmlEl.SetAttribute(localName, nsmgr.LookupNamespace(prefix), value);
            else
            {
                XmlAttribute a= xmlDoc.CreateAttribute(prefix, localName, nsmgr.LookupNamespace(prefix));
                a.Value = value;
                xmlEl.Attributes.Append(a);
            }
        }

        internal void RemoveAttribute(string name)
        {
            string prefix = null;
            string localName = name;
            if (name.Contains(":"))
            {
                prefix = name.Split(':')[0];
                localName = name.Split(':')[1];
            }

            if (xmlEl.HasAttribute(localName))
            {
                if (prefix != null)
                    xmlEl.RemoveAttribute(localName, nsmgr.LookupNamespace(prefix));
                else
                    xmlEl.RemoveAttribute(name);
            }
        }


        public T FindChildOrCreate<T>(INSERT_POS pos=INSERT_POS.LAST) where T : Node
        {
            T result= (T)childNodes.Where(x => x is T).FirstOrDefault();
            if(result==null)
                switch(pos)
                {
                    case INSERT_POS.LAST:
                        result = NewNodeLast<T>();
                        break;
                    case INSERT_POS.FIRST:
                        result = NewNodeFirst<T>();
                        break;
                    default:
                        throw new Exception($"Не реализовано для INSERT_POS.{pos.ToString()}");
                }
            return result;
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
                            result.Add(new GridCol(item, this, result.Where(x => x is GridCol).Count()));
                            break;
                        case "w:tr":
                            int trInd = result.Where(x => x is Tr).Count();
                            Tr newTr = new Tr(item, this, trInd);
                            result.Add(newTr);
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
                        case "w:ins":
                            result.Add(new Ins(item, this));
                            break;
                        case "w:pgMar":
                            result.Add(new PageMarginNode(item, this));
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
                        case "w:numPr":
                            result.Add(new NumPr(item, this));
                            break;
                        case "w:ilvl":
                            result.Add(new Ilvl(item, this));
                            break;
                        case "w:numId":
                            result.Add(new NumId(item, this));
                            break;
                      
                        case "w:abstractNum":
                            result.Add(new AbstractNum(item, this));
                            break;
                        case "w:lvl":
                            result.Add(new Lvl(item, this));
                            break;
                        case "w:start":
                            result.Add(new Start(item, this));
                            break;
                        case "w:lvlText":
                            result.Add(new LvlText(item, this));
                            break;
                        case "w:lvlJc":
                            result.Add(new LvlJc(item, this));
                            break;
                        case "w:num":
                            result.Add(new Num(item, this));
                            break;
                        case "w:abstractNumId":
                            result.Add(new AbstractNumId(item, this));
                            break;
                        case "w:nsid":
                            result.Add(new Nsid(item, this));
                            break;
                        case "w:multiLevelType":
                            result.Add(new MultiLevelType(item, this));
                            break;
                        case "w:tmpl":
                            result.Add(new Tmpl(item, this));
                            break;
                        case "w:rPrChange":
                            result.Add(new RprChange(item, this));
                            break;
                        case "w:pPrChange":
                            result.Add(new PprChange(item, this));
                            break;
                        case "w:tblPrChange":
                            result.Add(new TblPrChange(item, this));
                            break;
                        case "w:tcPrChange":
                            result.Add(new TcPrChange(item, this));
                            break;
                        case "w:tblGridChange":
                            result.Add(new TblGridChange(item, this));
                            break;
                        case "w:trPrChange":
                            result.Add(new TrPrChange(item, this));
                            break;
                        case "w:sectPrChange":
                            result.Add(new SectPrChange(item, this));
                            break;
                        case "w:sdt":
                            result.Add(new Sdt(item, this));
                            break;
                        case "w:sdtPr":
                            result.Add(new StdPr(item, this));
                            break;
                        case "w:docPartObj":
                            result.Add(new DocPartObj(item, this));
                            break;
                        case "w:docPartGallery":
                            result.Add(new DocPartGallery(item, this));
                            break;
                        case "w:fldChar":
                            result.Add(new FldChar(item, this));
                            break;
                        case "w:instrText":
                            result.Add(new InstrText(item, this));
                            break;
                        case "w:sdtContent":
                            result.Add(new SdtContent(item, this));
                            break;
                        case "w:footerReference":
                            result.Add(new FooterReference(item, this));
                            break;
                        case "Relationship":
                            result.Add(new Relationship(item, this));
                            break;
                        case "Override":
                            result.Add(new Override(item, this));
                            break;
                        case "w:jc":
                            result.Add(new Jc(item, this));
                            break;
                        case "w:id":
                            result.Add(new IdNode(item, this));
                            break;
                        case "w:customXmlInsRangeStart":
                            result.Add(new CustomXmlInsRangeStart(item, this));
                            break;
                        case "w:customXmlInsRangeEnd":
                            result.Add(new CustomXmlInsRangeEnd(item, this));
                            break;
                        case "w:headerReference":
                            result.Add(new HeaderReference(item, this));
                            break;
                        case "w:highlight":
                            result.Add(new Highlight(item, this));
                            break;
                        case "w:pgSz":
                            result.Add(new PgSz(item, this));
                            break;
                        case "w:pgNumType":
                            result.Add(new PgNumType(item, this));
                            break;
                        default:
                            result.Add(new Node(item, this, item.Name));
                            break;
                    }

                    if(result.Count() > 1)
                    { 
                        result[result.Count()-2].NextNode = result.Last();
                        result[result.Count() - 1].PrevNode = result[result.Count() - 2];
                    }

                }
                result.AddRange(baseStyleNodes);
                return result;
            }
        }

        internal List<Node> baseStyleNodes = new List<Node>();

        public virtual void CreateChangeNode(string changeNodeName = "w:pPrChange", XmlElement moveChangeNodeTo = null, string author = "TDV")
        {
            XmlElement oldNode = this.CopyXmlElement();
            if (moveChangeNodeTo == null)
                moveChangeNodeTo = xmlEl;
            XmlElement nChange = (XmlElement)moveChangeNodeTo.SelectSingleNode(changeNodeName, nsmgr);
            //создать ноду w: rPrChange если она не создана
            if (nChange == null)
            {
                nChange = xmlDoc.CreateElement(changeNodeName, xmlDoc.DocumentElement.NamespaceURI);
                nChange.SetAttribute("id", xmlEl.NamespaceURI, (xmlDoc.GetLastId() + 1).ToString());
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
                foreach (XmlElement forDel in sectChangeNode.SelectNodes("w:pgNumType", nsmgr))
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
            T result = Activator.CreateInstance<T>();
            result.xmlDoc = xmlEl.OwnerDocument;
            result.parent = this;
            result.nsmgr = nsmgr;
            result.InitXmlElement();
            return result;
        }

        public T NewNodeAfter<T>(XmlElement after) where T : Node
        {
            T result = NewNode<T>();
            xmlEl.InsertAfter(result.xmlEl, after);
            return result;
        }
        public T NewNodeAfter<T>(Node after) where T : Node
        {
            T result = NewNode<T>();
            xmlEl.InsertAfter(result.xmlEl, after.xmlEl);
            return result;
        }

        public T NewNodeBefore<T>(XmlElement before) where T : Node
        {
            T result = NewNode<T>();
            xmlEl.InsertBefore(result.xmlEl, before);
            return result;
        }
        public T NewNodeBefore<T>(Node before) where T : Node
        {
            T result = NewNode<T>();
            xmlEl.InsertBefore(result.xmlEl, before.xmlEl);
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
            xmlEl = xmlDoc.CreateElement(qualifiedName,xmlDoc.DocumentElement.NamespaceURI);
        }


        public virtual string Text
        {
            get
            {
                if(xmlEl!=null)
                    return xmlEl.InnerText;
                return null;
            }
            set
            {
                xmlEl.InnerText = value;
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