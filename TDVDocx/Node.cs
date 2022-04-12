using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml;
using System.Linq;
using System.Runtime.CompilerServices;
[assembly:InternalsVisibleTo("TDVDocx.Test")]


namespace TDV.Docx {
  public class ParagraphStyle {
    public ParagraphStyle(HORIZONTAL_ALIGN horizontalAlign,
        Border borderLeft,
        Border borderRight,
        Border borderTop,
        Border borderBottom,
        Border borderBetween,
        Border borderBar,
        Size spacingBefore,
        Size spacingAfter,
        double spacingLine,
        Size indentingFirstLine,
        Size indentingHanging,
        Size indentingLeft,
        Size indentingRight,
        int? numId = null,
        int numLevel = 0
        ) {
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
      this.indentingFirstLine = indentingFirstLine;
      this.indentingHanging = indentingHanging;
      this.indentingLeft = indentingLeft;
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
    public Size spacingBefore;
    public Size spacingAfter;
    public double spacingLine;
    public Size indentingFirstLine;
    public Size indentingHanging;
    public Size indentingLeft;
    public Size indentingRight;


    public ParagraphStyle Copy() {
      return new ParagraphStyle(horizontalAlign, borderLeft, borderRight, borderTop, borderBottom, borderBetween, borderBar, spacingBefore,
          spacingAfter, spacingLine, indentingFirstLine, indentingHanging, indentingLeft, indentingRight);
    }

    public ParagraphStyle SetHorizontalAlign(HORIZONTAL_ALIGN value) {
      horizontalAlign = value;
      return this;
    }
  }

  public class RunStyle {
    public static bool operator ==(RunStyle a, RunStyle b) {
      if (a is null && b is null)
        return true;
      if (!(a is null) && !(b is null)) {
        if (a.isBold == b.isBold &&
            a.font == b.font &&
            a.fontSize == b.fontSize &&
            a.isItalic == b.isItalic &&
            a.isStrike == b.isStrike &&
            a.underline == b.underline &&
            a.color == b.color &&
            a.highlight == b.highlight &&
            a.border == b.border)
          return true;
        else
          return false;
      }
      else
        return false;
    }

    public RunStyle SetIsBold(bool value) {
      isBold = value;
      return this;
    }
    public static bool operator !=(RunStyle a, RunStyle b) {

      return !(a == b);
    }
    public RunStyle(bool isBold,
        string font,
        double? fontSize,
        bool isItalic,
        bool isStrike,
        LINE_TYPE underline,
        string color,
        string highlight,
        Border border) {
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
    public double? fontSize;
    public bool isItalic;
    public bool isStrike;
    public LINE_TYPE underline;
    public string color;
    public string highlight;
    public Border border;
    public RunStyle Copy() {
      return new RunStyle(isBold, font, fontSize, isItalic, isStrike, underline, color, highlight, border);
    }

  }

  public enum HORIZONTAL_ALIGN { LEFT, CENTER, RIGHT, BOTH }

  public enum BORDER_TYPE {
    LEFT, RIGHT, TOP, BOTTOM,
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

  /// <summary>
  /// тип связи. для заголовков\футеров
  /// </summary>
  public enum REFERENCE_TYPE { FIRST, EVEN, DEFAULT }

  public enum INSERT_POS { FIRST, LAST }
  public class Node {
    protected Node(string qualifiedName = "") {
      this.qualifiedName = qualifiedName;
      disableNodeChanged = false;
    }

    public static bool operator ==(Node a, Node b) {
      if (a is null && b is null)
        return true;
      if (a is null || b is null)
        return false;
      return a.XmlEl == b.XmlEl;
    }
    public static bool operator !=(Node a, Node b) {
      return !(a == b);
    }
    protected bool disableNodeChanged;
    public Section Section {
      get {
        foreach (Section s in GetDocxDocument().Document.Sections)
          if (s.ChildNodes.Where(x => x.XmlEl == XmlEl).Any())
            return s;
        return null;
      }
    }
    public Node NextNode {
      get {
        if (Parent is null)
          return null;
        List<Node> childNodes = Parent.ChildNodes;
        Node currNode = childNodes.Where(x => x.XmlEl == XmlEl).First();
        int currNodeIndex = childNodes.IndexOf(currNode);
        if (childNodes.Count() > currNodeIndex + 1) {
          return childNodes[currNodeIndex + 1];
        }
        else {
          return null;
        }
      }
    }

    public Node NextNodeRecurcieve {
      get {
        if (ChildNodes.Count > 0)
          return ChildNodes.First();
        else if (NextNode != null)
          return NextNode;
        else if (NextNode == null)
          return Parent.NextNode;
        else
          return null;
      }
    }

    public Node PrevNode {
      get {
        if (Parent is null)
          return null;
        List<Node> childNodes = Parent.ChildNodes;
        Node currNode = childNodes.Where(x => x.XmlEl == XmlEl).First();
        int currNodeIndex = childNodes.IndexOf(currNode);
        if (currNodeIndex - 1 >= 0) {
          return childNodes[currNodeIndex - 1];
        }
        else {
          return null;
        }
      }
    }

    public virtual void ApplyAllFixes() {
      foreach (Node n in ChildNodes) {
        n.ApplyAllFixes();
      }
    }

    public string GenerateGuid() {
      return Guid.NewGuid().ToString().Substring(0, 8).ToUpper();
    }

    public int GenerateId(int len = 9) {
      return new Random().Next((int)Math.Pow(10.0, (double)(len - 1)), (int)Math.Pow(10.0, (double)len) - 1);
    }

    public XmlElement CopyXmlElement() {
      XmlElement result = XmlDoc.CreateElement(XmlEl.Prefix, XmlEl.LocalName, XmlEl.NamespaceURI);
      result.InnerXml = XmlEl.InnerXml;
      foreach (XmlAttribute att in XmlEl.Attributes)
        result.SetAttribute(att.LocalName, att.NamespaceURI, att.Value);
      return result;
    }

    public Node(Node parent, string qualifiedName = "") {
      this.Parent = parent;
      this.qualifiedName = qualifiedName;
      this.Nsmgr = parent.Nsmgr;
      this.XmlDoc = parent.XmlDoc;
      InitXmlElement();
      if (this is PProp || this is RProp)
        parent.XmlEl.InsertBefore(XmlEl, parent.XmlEl.FirstChild);
      else
        parent.XmlEl.AppendChild(XmlEl);
    }
    public Node(XmlElement xmlElement, Node parent, string qualifiedName) : this(qualifiedName) {
      this.Parent = parent;
      XmlDoc = xmlElement.OwnerDocument;
      Nsmgr = parent.Nsmgr;
      XmlEl = xmlElement;
    }
    public XmlDocument XmlDoc;
    public XmlNamespaceManager Nsmgr;
    public Node Parent;

    internal string qualifiedName;

    public virtual void NodeChanded() {
    }

    public T GetParentRecurcieve<T>() where T : Node {
      if (Parent == null)
        return null;
      if (Parent is T)
        return (T)Parent;
      else
        return Parent.GetParentRecurcieve<T>();
    }

    public T FindChild<T>() where T : Node {
      return (T)ChildNodes.Where(x => x is T).FirstOrDefault();
    }

    internal string GetAttribute(string name) {
      string prefix = null;
      string localName = name;
      if (name.Contains(":")) {
        prefix = name.Split(':')[0];
        localName = name.Split(':')[1];
      }
      foreach (XmlAttribute a in XmlEl.Attributes) {
        if (a.Name == name || (a.LocalName == localName && a.Prefix == prefix))
          return a.Value;
      }
      throw new KeyNotFoundException();
    }
    internal void SetAttribute(string name, string value) {
      string prefix = null;
      string localName = name;
      if (name.Contains(":")) {
        prefix = name.Split(':')[0];
        localName = name.Split(':')[1];
      }

      if (HasAttribute(name))
        XmlEl.SetAttribute(localName, Nsmgr.LookupNamespace(prefix), value);
      else {
        XmlAttribute a = XmlDoc.CreateAttribute(prefix, localName, Nsmgr.LookupNamespace(prefix));
        a.Value = value;
        XmlEl.Attributes.Append(a);
      }
    }

    internal bool HasAttribute(string name) {
      string prefix = null;
      string localName = name;
      if (name.Contains(":")) {
        prefix = name.Split(':')[0];
        localName = name.Split(':')[1];
      }
      if (string.IsNullOrEmpty(prefix))
        return XmlEl.HasAttribute(localName);
      else
        return XmlEl.HasAttribute(localName, Nsmgr.LookupNamespace(prefix));
    }

    internal void RemoveAttribute(string name) {
      string prefix = null;
      string localName = name;
      if (name.Contains(":")) {
        prefix = name.Split(':')[0];
        localName = name.Split(':')[1];
      }

      if (XmlEl.HasAttribute(localName) || XmlEl.HasAttribute(localName, Nsmgr.LookupNamespace(prefix))) {
        if (prefix != null)
          XmlEl.RemoveAttribute(localName, Nsmgr.LookupNamespace(prefix));
        else
          XmlEl.RemoveAttribute(name);
      }
    }

    public T FindChildOrCreate<T>(INSERT_POS pos = INSERT_POS.LAST) where T : Node {
      T result = (T)ChildNodes.Where(x => x is T).FirstOrDefault();
      if (result == null)
        switch (pos) {
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

    public List<T> FindChilds<T>() where T : Node {
      return ChildNodes.Where(x => x is T).Select(x => (T)x).ToList();
    }

    public List<T> FindChildsRecurcieve<T>() where T : Node {
      List<T> result = new List<T>();
      result = ChildNodes.Where(x => x is T).Select(x => (T)x).ToList();
      foreach (Node child in ChildNodes) {
        result.AddRange(child.FindChildsRecurcieve<T>());
      }
      return result;
    }

    public virtual List<Node> ChildNodes {
      get {
        List<Node> result = new List<Node>();
        if (XmlEl == null)
          return result;
        foreach (var el in XmlEl.ChildNodes) {
          if (!(el is XmlElement))
            continue;
          XmlElement item = (XmlElement)el;
          switch (item.Name) {
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
              Tc newTc = new Tc(item, this, colInd);
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
            case "w:hyperlink":
              result.Add(new Hyperlink(item, this));
              break;
            case "w:semiHidden":
              result.Add(new SemiHidden(item, this));
              break;
            case "w:unhideWhenUsed":
              result.Add(new UnhideWhenUsed(item, this));
              break;
            case "w:name":
              result.Add(new Name(item, this));
              break;
            case "w:basedOn":
              result.Add(new BasedOn(item, this));
              break;
            case "w:uiPriority":
              result.Add(new UiPriority(item, this));
              break;
            case "w:rStyle":
              result.Add(new RStyleNode(item, this));
              break;
            case "w:noProof":
              result.Add(new NoProof(item, this));
              break;
            case "w:t":
              result.Add(new T(item, this));
              break;
            case "w:rsids":
              result.Add(new Rsids(item, this));
              break;
            case "w:rsidRoot":
              result.Add(new RsidRoot(item, this));
              break;
            case "w:rsid":
              result.Add(new Rsid(item, this));
              break;
            case "w:lang":
              result.Add(new Lang(item, this));
              break;
            case "w:commentRangeStart":
              result.Add(new CommentRangeStart(item, this));
              break;
            case "w:commentRangeEnd":
              result.Add(new CommentRangeEnd(item, this));
              break;
            case "w:comment":
              result.Add(new Comment(item, this));
              break;
            case "w:commentReference":
              result.Add(new CommentReference(item, this));
              break;
            case "w:delText":
              result.Add(new DelText(item, this));
              break;
            case "w:sz":
              result.Add(new Sz(item, this));
              break;
            case "w:szCs":
              result.Add(new SzCs(item, this));
              break;
            case "w:pStyle":
              result.Add(new PStyle(item, this));
              break;
            case "w:rFonts":
              result.Add(new RFonts(item, this));
              break;
            case "w:b":
              result.Add(new B(item, this));
              break;
            case "w:i":
              result.Add(new I(item, this));
              break;
            case "w:strike":
              result.Add(new Strike(item, this));
              break;
            case "w:pict":
              result.Add(new Pict(item, this));
              break;
            case "w:titlePg":
              result.Add(new TitlePg(item, this));
              break;
            case "a:themeElements":
              result.Add(new ThemeElements(item, this));
              break;
            case "a:fontScheme":
              result.Add(new FontScheme(item, this));
              break;
            case "a:majorFont":
              result.Add(new MajorFont(item, this));
              break;
            case "a:latin":
              result.Add(new Latin(item, this));
              break;
            case "a:objectDefaults":
              result.Add(new ObjectDefaults(item, this));
              break;
            case "w:top":
              result.Add(new Top(item, this));
              break;
            case "w:left":
              result.Add(new Left(item, this));
              break;
            case "w:bottom":
              result.Add(new Bottom(item, this));
              break;
            case "w:right":
              result.Add(new Right(item, this));
              break;
            case "w:insideH":
              result.Add(new InsideH(item, this));
              break;
            case "w:insideV":
              result.Add(new InsideV(item, this));
              break;
            case "a:clrScheme":
              result.Add(new ClrScheme(item, this));
              break;
            case "a:srgbClr":
              result.Add(new SrgbClr(item, this));
              break;
            case "w:proofErr":
              result.Add(new ProofErr(item, this));
              break;
            case "w:bookmarkEnd":
              result.Add(new BookmarkEnd(item, this));
              break;
            case "w:bookmarkStart":
              result.Add(new BookmarkStart(item, this));
              break;
            case "w:br":
              result.Add(new Br(item, this));
              break;
            case "w:color":
              result.Add(new WColor(item, this));
              break;
            case "w:u":
              result.Add(new U(item, this));
              break;
            case "w:pPrDefault":
              result.Add(new PPrDefault(item, this));
              break;
            case "w:rPrDefault":
              result.Add(new RPrDefault(item, this));
              break;
            case "w:docDefaults":
              result.Add(new DocDefaults(item, this));
              break;
            case "w:cols":
              result.Add(new Cols(item, this));
              break;
            case "w:type":
              result.Add(new WType(item, this));
              break;
            case "m:oMathPara":
              result.Add(new OMathPara(item, this));
              break;
            case "m:oMath":
              result.Add(new OMath(item, this));
              break;
            case "m:oMathParaPr":
              result.Add(new OMathParaPr(item, this));
              break;
            case "m:jc":
              result.Add(new MJc(item, this));
              break;
            default:
              result.Add(new Node(item, this, item.Name));
              break;
          }

          if (result.Count() > 1) {
            //result[result.Count()-2].NextNode = result.Last();
            //result[result.Count() - 1].PrevNode = result[result.Count() - 2];
          }

        }
        result.AddRange(baseStyleNodes);
        return result;
      }
    }

    internal List<Node> baseStyleNodes = new List<Node>();
    private bool ChangeNodeCreated = false;
    public virtual void CreateChangeNode<T>(string author) where T : ChangeNode {
      if (ChangeNodeCreated)
        return;
      if (Parent.FindChild<T>() != null || FindChild<T>() != null)
        return;
      //return (T) FindChild<T>();
      ChangeNode changeNode = NewNodeLast<T>();
      if (changeNode == null)
        throw new NotImplementedException();
      ChangeNodeCreated = true;
      changeNode.Author = author;
    }


    public XmlElement XmlEl;

    public virtual void Delete() {
      if (Parent != null && Parent.XmlEl.SelectSingleNode(XmlEl.Name, Nsmgr) != null) {
        Parent.XmlEl.RemoveChild(XmlEl);
      }
    }

    public void MoveTo(Node nodeTo) {
      Delete();
      if (nodeTo.Parent != this) {
        nodeTo.XmlEl.AppendChild(XmlEl);
        Parent = nodeTo;
      }
    }

    public void MoveAfter(Node after) {
      Node nodeTo = after.Parent;
      Delete();
      nodeTo.XmlEl.InsertAfter(XmlEl, after.XmlEl);
      Parent = after.Parent;
    }
    public void MoveBefore(Node before) {
      Node nodeTo = before.Parent;
      Delete();
      nodeTo.XmlEl.InsertBefore(XmlEl, before.XmlEl);
      Parent = before.Parent;
    }

    public void Clear() {
      XmlEl.RemoveAll();
      ChildNodes.Clear();
    }

    //добавляет новую НОДУ в конец списка
    private T NewNode<T>() where T : Node {
      T result = Activator.CreateInstance<T>();
      result.XmlDoc = XmlEl.OwnerDocument;
      result.Parent = this;
      result.Nsmgr = Nsmgr;
      result.InitXmlElement();
      return result;
    }

    public T NewNodeAfter<T>(XmlElement after) where T : Node {
      T result = NewNode<T>();
      XmlEl.InsertAfter(result.XmlEl, after);
      return result;
    }
    public T NewNodeAfter<T>(Node after) where T : Node {
      T result = NewNode<T>();
      XmlEl.InsertAfter(result.XmlEl, after.XmlEl);
      return result;
    }

    public T NewNodeBefore<T>(XmlElement before) where T : Node {
      T result = NewNode<T>();
      XmlEl.InsertBefore(result.XmlEl, before);
      return result;
    }
    public T NewNodeBefore<T>(Node before) where T : Node {
      T result = NewNode<T>();
      XmlEl.InsertBefore(result.XmlEl, before.XmlEl);
      return result;
    }

    public T NewNodeFirst<T>() where T : Node {
      T result = NewNode<T>();
      XmlEl.InsertBefore(result.XmlEl, XmlEl.FirstChild);
      return result;
    }

    public T NewNodeLast<T>() where T : Node {
      T result = NewNode<T>();
      XmlEl.AppendChild(result.XmlEl);
      return result;
    }

    /// <summary>
    /// Создает новый XmlElement. Необходимо переопределеять в классах наследниках
    /// </summary>
    public virtual void InitXmlElement() {
      XmlEl = XmlDoc.CreateElement(qualifiedName, XmlDoc.DocumentElement.NamespaceURI);
    }

    public virtual string Text {
      get {
        if (XmlEl != null)
          return XmlEl.InnerText;
        return null;
      }
      set {
        XmlEl.InnerText = value;
      }
    }

    public DocxDocument GetDocxDocument() {
      DocxDocument result = null;

      if (this is BaseNode)
        result = ((BaseNode)this).DocxDocument;
      else if (Parent != null)
        result = Parent.GetDocxDocument();
      return result;
    }

    public virtual int GetWordsCnt() {
      return Text?.Trim().Split(' ').Count() ?? 0;
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="wordsLimit">Максимальное количество слов</param>
    /// <param name="nodesLimit">Максимально количество нод</param>
    /// <param name="calcEmptyParagrapsh"></param>
    /// <returns></returns>
    public List<Node> GetNextNodesList(int? wordsLimit = null, int? nodesLimit = null, bool calcEmptyNodes = true) {
      List<Node> result = new List<Node>();
      Node currNode = NextNode;
      int wordsCnt = 0;
      while (currNode != null) {
        if (nodesLimit.HasValue && result.Count == nodesLimit) {
          break;
        }

        if (wordsLimit.HasValue) {
          wordsCnt += currNode.GetWordsCnt();
          if (wordsLimit <= wordsCnt) {
            break;
          }
        }

        if (calcEmptyNodes || !string.IsNullOrEmpty(currNode.Text)) {
          result.Add(currNode);
        }
        currNode = currNode.NextNode;
      }
      return result;
    }

    public List<Node> GetPrevNodesList(int? wordsLimit = null, int? nodesLimit = null, int? symbolsLimit = null, bool calcEmptyNodes = true) {
      List<Node> result = new List<Node>();
      Node currNode = PrevNode;
      int wordsCnt = 0;
      int symbolsCnt = 0;
      while (currNode != null) {
        if (nodesLimit.HasValue && result.Count == nodesLimit) {
          break;
        }

        if (wordsLimit.HasValue) {

          if (wordsLimit < wordsCnt + currNode.GetWordsCnt()) {
            break;
          }
          wordsCnt += currNode.GetWordsCnt();
        }

        if (symbolsLimit.HasValue) {
          if (symbolsLimit < symbolsCnt + currNode.Text.Length) {
            break;
          }
          symbolsLimit += currNode.Text.Length;
        }

        if (calcEmptyNodes || !string.IsNullOrEmpty(currNode.Text)) {
          result.Add(currNode);
        }
        currNode = currNode.PrevNode;
      }

      result.Reverse();
      return result;
    }

    public T Cast<T>() where T : Node {
      return (T)this;
    }
  }

  public interface ICorrectable {
    /// <summary>
    /// устанавливает режим правки
    /// mode = "del", "ins"
    /// </summary>
    void SetCorrectionMode(string mode, string author = "TDV");
  }
}