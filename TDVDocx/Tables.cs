using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms.VisualStyles;
using System.Xml;

namespace TDV.Docx {
  public class CellStyle {
    public CellStyle(
        VERTICAL_ALIGN vAlign,
                Border borderLeft,
                Border borderRight,
                Border borderTop,
                Border borderBottom,
                Size width
        ) {
      this.vAlign = vAlign;
      this.borderLeft = borderLeft;
      this.borderRight = borderRight;
      this.borderTop = borderTop;
      this.borderBottom = borderBottom;
      this.width = width;
    }

    public VERTICAL_ALIGN vAlign;
    public Border borderLeft;
    public Border borderRight;
    public Border borderTop;
    public Border borderBottom;
    public Size width;
  }

  public class TableStyle {
    public TableStyle(
        Border borderLeft,
    Border borderRight,
    Border borderTop,
    Border borderBottom,
    Border borderInsideH,
    Border borderInsideV,
    bool applyBorderToCells,
    Size width,
    Size indentingWidth
    ) {
      this.BorderLeft = borderLeft;
      this.BorderRight = borderRight;
      this.BorderTop = borderTop;
      this.BorderBottom = borderBottom;
      this.BorderInsideH = borderInsideH;
      this.BorderInsideV = borderInsideV;
      this.ApplyBorderToCells = applyBorderToCells;
      this.Width = width;
      this.IndentingWidth = indentingWidth;
    }

    public Border BorderLeft;
    public Border BorderRight;
    public Border BorderTop;
    public Border BorderBottom;
    public Border BorderInsideH;
    public Border BorderInsideV;
    public bool ApplyBorderToCells;
    public Size Width;
    public Size IndentingWidth;
  }
  public class Table : Node {
    public Table() : base("w:tbl") { }
    public Table(Node parent) : base(parent, "w:tbl") { }
    public Table(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:tbl") { }

    public TableStyle GetTableStyle() {
      return new TableStyle(borderLeft: BorderLeft, BorderRight, BorderTop, BorderBottom, BorderInsideH, BorderInsideV, true, Width, IndentingWidth);
    }

    #region Границы
    public Border BorderLeft {
      get { return TblPr.TblBorders.Left; }
      set { TblPr.TblBorders.Left = value; }
    }

    public Border BorderRight {
      get { return TblPr.TblBorders.Right; }
      set { TblPr.TblBorders.Right = value; }
    }

    public Border BorderBottom {
      get { return TblPr.TblBorders.Bottom; }
      set { TblPr.TblBorders.Bottom = value; }
    }

    public Border BorderTop {
      get { return TblPr.TblBorders.Top; }
      set { TblPr.TblBorders.Top = value; }
    }

    public Border BorderInsideH {
      get { return TblPr.TblBorders.InsideH; }
      set { TblPr.TblBorders.InsideH = value; }
    }

    public Border BorderInsideV {
      get { return TblPr.TblBorders.InsideV; }
      set { TblPr.TblBorders.InsideV = value; }
    }
    #endregion

    public Size Width {
      get {
        return TblPr.TblW.Width;
      }
      set {
        TblPr.TblW.Width = value;
      }
    }

    public Size IndentingWidth {
      get {
        return TblPr.TblInd.Width;
      }
      set {
        TblPr.TblInd.Width = value;
      }
    }

    /// <summary>
    /// пробегает по всем столбцам и сравнивает значение шириры с tblGrid\gridCol
    /// </summary>
    public void FixColumnsSizes(string author = "TDV") {

      //убрать лишние gridCols. иногда их больше чем по факту колонок в таблице
      //Вычисляю к-во максимальное к-во колонок
      int maxCellCnt = 0;

      foreach (Tr row in Rows)
        if (maxCellCnt < row.Cells.Count())
          maxCellCnt = row.Cells.Count();

      //Вычислить медианные значения ширины колонок
      List<List<int>> cellSizes = new List<List<int>>();

      for (int cellIndex = 0; cellIndex < maxCellCnt; cellIndex++) {
        List<int> cellSizeList = new List<int>();
        for (int rowIndex = 0; rowIndex < Rows.Count; rowIndex++) {
          Tc cell = GetCell(rowIndex, cellIndex);
          if (cell == null)
            continue;
          cellSizeList.Add(cell.Width.ValuePoints);
        }
        cellSizes.Add(cellSizeList);
      }

      //обновляю значения ширины колонок TblGrid.GridCols
      for (int i = 0; i < cellSizes.Count; i++) {
        TblGrid.GridCols[i].Width = new Size(cellSizes[i].Median());
      }

      //удаляю лишние колонки
      while (maxCellCnt < TblGrid.GridCols.Count)
        TblGrid.GridCols.Last().Delete();

      List<GridCol> gridCols = TblGrid.GridCols;
      foreach (Tr row in Rows) {
        //иногда в ячейках строки бывают лишние colspan. например к-во ячееек == 
        if (row.Cells.Count == maxCellCnt) {
          foreach (Tc cell in row.Cells)
            if (cell.GridSpan != 1)
              cell.GridSpan = 1;
        }

        if (row.Cells.Count == 1 && row.Cells.First().GridSpan != 1) {
          row.Cells.First().GridSpan = maxCellCnt;
        }

        for (int colInd = 0; colInd < row.Cells.Count(); colInd++) {
          Tc cell = row.Cells[colInd];

          if (cell.ColSpan <= 1) {
            cell.CompareWidth(gridCols[colInd].Width, author);
          }
          else {
            //вычислить общую ширину для объединенных ячеек
            Size newCellSize = new Size(0);
            for (int i = colInd; i < colInd + cell.ColSpan; i++) {
              if (i < gridCols.Count - 1) {
                newCellSize = newCellSize + gridCols[i].Width;
              }
            }
            cell.CompareWidth(newCellSize, author);
          }
        }
      }
    }

    /// <summary>
    /// Создает ChangeNode  для tblPr и  tblGrid
    /// </summary>
    internal void CreateChangeNodes(string author = "TDV") {
      if (TblPr.FindChild<TblPrChange>() == null) {
        TblPr.CreateChangeNode<TblPrChange>(author);
        TblGrid.CreateChangeNode<TblGridChange>(author);
        foreach (Tr row in Rows) {
          row.trPr.CreateChangeNode<TrPrChange>(author);
          foreach (Tc cell in row.Cells)
            cell.TcProp.CreateChangeNode<TcPrChange>(author);
        }
      }
    }
    public override void ApplyAllFixes() {

      FindChild<TableProp>()?.FindChild<TblPrChange>()?.Delete();
      FindChild<TableGrid>()?.FindChild<TblGridChange>()?.Delete();

      foreach (Tr tr in Rows) {
        if (tr.FindChild<TrProp>()?.FindChild<Del>() != null) {
          tr.Delete();
          continue;
        }

        tr.FindChild<TrProp>()?.FindChild<Ins>()?.Delete();
        tr.FindChild<TrProp>()?.FindChild<TrPrChange>()?.Delete();

        foreach (Tc tc in tr.Cells) {
          tc.FindChild<TcProp>()?.FindChild<TcPrChange>()?.Delete();
          foreach (Node n in tc.ChildNodes) {
            if (n is Paragraph) {
              ((Paragraph)n).ApplyAllFixes();
            }
            else if (n is Table) {
              ((Table)n).ApplyAllFixes();
            }
          }
        }
      }
    }

    public void CompareStyle(TableStyle style, string author = "TDV") {
      CompareBorders(style.BorderLeft, style.BorderRight, style.BorderTop, style.BorderBottom,
          style.BorderInsideH, style.BorderInsideV, style.ApplyBorderToCells, author);
      CompareIndenting(style.IndentingWidth, author);
      CompareTblWidth(style.Width, author);
    }
    public TableProp TblPr {
      get {
        return FindChildOrCreate<TableProp>();
      }
    }
    public void CompareBorders(Border left, Border right, Border top, Border bottom, Border insideH, Border insideV, bool applyToCells, string author = "TDV") {
      TblPr.TblBorders.CompareBorders(top, bottom, left, right, insideH, insideV, author);
      foreach (Tr tr in Rows) {
        foreach (Tc tc in tr.Cells) {
          if (applyToCells) {
            tc.CompareBorders(left, right, top, bottom, false, author);
          }
        }
      }
    }

    public void CompareIndenting(Size indenting, string author) {
      TblPr.CompareIndenting(indenting, author);
    }

    public void CompareTblWidth(Size width, string author) {
      TblPr.CompareTblWidth(width, author);
    }

    public TableGrid TblGrid {
      get {
        return FindChildOrCreate<TableGrid>();
      }
    }

    public Tc GetCell(int row, int col) {
      if (Rows.Count - 1 < row)
        return null;
      if (Rows[row].Cells.Count - 1 < col)
        return null;
      return Rows[row].Cells[col];
    }

    public List<Tr> Rows {
      get { return FindChilds<Tr>(); }
    }

    public override string ToString() {
      string result = "<Table> ";
      foreach (Tr row in Rows) {
        foreach (Tc cell in row.Cells) {
          foreach (Paragraph p in cell.Paragraphs) {
            result += p.ToString();
          }
        }
      }
      return result;
    }
  }

  public class TableProp : Node {
    public TableProp() : base("w:tblPr") { }
    public TableProp(Node parent) : base(parent, "w:tblPr") { }
    public TableProp(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:tblPr") { }

    public TblBorders TblBorders {
      get {
        return FindChildOrCreate<TblBorders>();
      }
    }

    public Style Style {
      get {
        return FindChild<TblStyle>()?.Style ?? null;
      }
      set {
        FindChildOrCreate<TblStyle>().Style = value;
      }
    }

    public void CompareIndenting(Size width, string author) {
      if (TblInd.Width != width) {
        Table tbl = GetParentRecurcieve<Table>();
        tbl.CreateChangeNodes(author);
        TblInd.Width = width;
      }
    }

    public void CompareTblWidth(Size width, string author) {
      if (TblW.Width != width) {
        Table tbl = GetParentRecurcieve<Table>();
        tbl.CreateChangeNodes(author);
        TblW.Width = width;
      }
    }

    public TableWidth TblW {
      get {
        return FindChildOrCreate<TableWidth>();
      }
    }

    public TblInd TblInd {
      get {
        return FindChildOrCreate<TblInd>();
      }
    }

    public TblStyle TblStyle {
      get {
        return FindChildOrCreate<TblStyle>();
      }
    }
  }

  public class TblStyle : Node {
    public TblStyle() : base("w:tblStyle") { }
    public TblStyle(Node parent) : base(parent, "w:tblStyle") { }
    public TblStyle(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:tblStyle") { }

    public string StyleId {
      get { return XmlEl.GetAttribute("w:val"); }
    }

    public string Value {
      get {
        try {
          return GetAttribute("w:val");
        }
        catch (KeyNotFoundException) {
          return null;
        }
      }
      set {
        if (string.IsNullOrEmpty(value))
          RemoveAttribute("w:val");
        else
          SetAttribute("w:val", value);
      }
    }

    public Style Style {
      get {
        if (string.IsNullOrEmpty(Value))
          return null;
        DocxDocument docx = GetDocxDocument();
        return docx.Styles.GetStyleById(Value);
      }
      set {
        Value = value.StyleId;
      }
    }

  }
  public enum TABLE_WIDTH_TYPE {
    /// <summary>
    /// Ширина определяется общим алгоритмом компановки таблицы
    /// </summary>
    AUTO,
    /// <summary>
    /// Значение в 1/1440 дюйма или 1/567 см
    /// </summary>
    DXA,
    /// <summary>
    /// Нулевое значение
    /// </summary>
    NIL,
    /// <summary>
    /// Значение в процентах от ширины таблицы. 5000 - это 100%
    /// </summary>
    PCT
  }
  public class TableWidth : Node {
    public TableWidth() : base("w:tblW") { }
    public TableWidth(Node parent) : base(parent, "w:tblW") { }
    public TableWidth(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:tblW") { }

    public Size Width {
      get {
        Size parentSize = null;
        if (SizeType == TABLE_WIDTH_TYPE.PCT)
          parentSize = Section?.SectProp.WorkspaceWidth;
        return new Size(Int32.Parse(GetAttribute("w:w")), SizeType, parentSize);
      }
      set {
        SizeType = value.SizeType;
        SetAttribute("w:w", value.Value.ToString());
      }
    }

    public TABLE_WIDTH_TYPE SizeType {
      get {
        return (TABLE_WIDTH_TYPE)Enum.Parse(typeof(TABLE_WIDTH_TYPE), XmlEl.GetAttribute("w:type"), true);
      }
      set { XmlEl.SetAttribute("type", XmlEl.NamespaceURI, value.ToString().ToLower()); }
    }
  }

  /// <summary>
  /// Отступ таблицы (слева)
  /// </summary>
  public class TblInd : Node {
    public TblInd() : base("w:tblInd") { }
    public TblInd(Node parent) : base(parent, "w:tblInd") { }
    public TblInd(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:tblInd") { }

    public Size Width {
      get {
        try {

          return new Size(Int32.Parse(GetAttribute("w:w")), SizeType);
        }
        catch {
          return new Size(0);
        }
      }
      set {
        SetAttribute("w:w", value.ValuePoints.ToString());
        SizeType = value.SizeType;
      }
    }

    public TABLE_WIDTH_TYPE SizeType {
      get {
        if (!XmlEl.HasAttribute("w:type"))
          return TABLE_WIDTH_TYPE.NIL;
        return (TABLE_WIDTH_TYPE)Enum.Parse(typeof(TABLE_WIDTH_TYPE), XmlEl.GetAttribute("w:type"), true);
      }
      set { SetAttribute("w:type", value.ToString().ToLower()); }
    }
  }

  public class TableGrid : Node {
    public TableGrid() : base("w:tblGrid") { }
    public TableGrid(Node parent) : base(parent, "w:tblGrid") { }
    public TableGrid(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:tblGrid") { }

    public List<GridCol> GridCols {
      get { return FindChilds<GridCol>(); }
    }
  }

  public class GridCol : Node {
    public GridCol() : base("w:gridCol") { }
    public GridCol(Node parent) : base(parent, "w:gridCol") { }
    public GridCol(XmlElement xmlElement, Node parent, int colIndex) : base(xmlElement, parent, "w:gridCol") { ColIndex = colIndex; }

    private Table _parentTable;
    public Table ParentTable {
      get {
        if (_parentTable == null)
          _parentTable = GetParentRecurcieve<Table>();
        return _parentTable;
      }
    }

    public Size Width {
      get {
        return new Size(Int32.Parse(XmlEl.GetAttribute("w:w")), SizeType);
      }
      set {
        SetAttribute("w:w", value.ValuePoints.ToString());
        SizeType = value.SizeType;
      }
    }

    public readonly int ColIndex;
    public TABLE_WIDTH_TYPE SizeType {
      get {
        TABLE_WIDTH_TYPE result = TABLE_WIDTH_TYPE.AUTO;
        Enum.TryParse<TABLE_WIDTH_TYPE>(XmlEl.GetAttribute("w:type"), true, out result);
        return result;
      }
      set { XmlEl.SetAttribute("type", XmlEl.NamespaceURI, value.ToString().ToLower()); }
    }
    /// <summary>
    /// Задает значения ширины таблицы в режиме правки
    /// </summary>
    /// <param name="width">Значение в нормальных процентах (от 0 до 100). будет переведено в диапазон от 0 до 5000</param>
    /// <param name="type"></param>
    /// <param name="applyToColumns">Применить к столбцам таблицы</param>
    /// <param name="author"></param>
    public void CompareWidth(Size width, bool applyToColumns, string author = "TDV") {
      if (width != Width) {
        ParentTable.CreateChangeNodes(author);
        Width = width;
        if (applyToColumns)
          foreach (Tr row in ParentTable.Rows)
            row.Cells[this.ColIndex].CompareWidth(new Size(5000, TABLE_WIDTH_TYPE.PCT, ParentTable.Width), author);
      }
    }
  }

  public class Tr : Node {
    public Tr() : base("w:tr") { }
    public Tr(Node parent, int rowIndex) : base(parent, "w:tr") { RowIndex = rowIndex; }

    public Tr(XmlElement xmlElement, Node parent, int rowIndex) : base(xmlElement, parent, "w:tr") { RowIndex = rowIndex; }

    public TrProp trPr {
      get {
        var result = ChildNodes.Where(x => x is TrProp).Select(x => x).FirstOrDefault();
        if (result == null)
          result = new TrProp(this);
        return (TrProp)result;
      }
    }

    public readonly int RowIndex;

    public List<Tc> Cells {
      get { return ChildNodes.Where(x => x is Tc).Select(x => (Tc)x).ToList(); }
    }

    public void CorrectDel(string author = "TDV") {
      trPr.SetCorrectionMode("del", author);
      foreach (Tc cell in Cells)
        cell.CorrectDel(author);
    }


    public void CompareHeigth(int heigth, string author = "TDV") {
      trPr.CompareHeigth(heigth, author);
    }

    public override string ToString() {
      return Text;
    }
  }

  public class TrProp : Node, ICorrectable {
    public TrProp() : base("w:trPr") { }
    public TrProp(Node parent) : base(parent, "w:trPr") { }
    public TrProp(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:trPr") { }

    private Table _parentTable;
    public Table ParentTable {
      get {
        if (_parentTable == null)
          _parentTable = GetParentRecurcieve<Table>();
        return _parentTable;
      }
    }

    /// <summary>
    /// Изменение высоты строки в режиме правки
    /// </summary>
    /// <param name="heigth">-1: тег будет удален</param>
    /// <param name="author"></param>
    public void CompareHeigth(int heigth, string author) {
      if (TrHeight.Value != heigth) {
        //CreateChangeNode<TcPrChange>(author);
        ParentTable.CreateChangeNodes(author);
        TrHeight.Value = heigth;
      }
    }

    public TrHeight TrHeight {
      get {
        return FindChildOrCreate<TrHeight>();
      }
    }

    /// <summary>
    /// устанавливает режим правки
    /// mode = "del", "ins"
    /// </summary>
    public void SetCorrectionMode(string mode, string author = "TDV") {
      var cNode = (XmlElement)XmlEl.SelectSingleNode($"w:{mode}", Nsmgr);
      if (cNode == null) {
        cNode = (XmlElement)XmlDoc.CreateElement($"w:{mode}", XmlEl.NamespaceURI);
        cNode.SetAttribute("id", XmlEl.NamespaceURI, (GetDocxDocument().Document.GetNextId()).ToString());
        XmlEl.InsertBefore(cNode, XmlEl.FirstChild);
      }
      cNode.SetAttribute("author", XmlEl.NamespaceURI, author);
      cNode.SetAttribute("date", XmlEl.NamespaceURI, DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ssZ"));
    }
  }

  public class TrHeight : Node {
    public TrHeight() : base("w:trHeight") { }
    public TrHeight(Node parent) : base(parent, "w:trHeight") { }
    public TrHeight(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:trHeight") { }

    /// <summary>
    /// При установке значения -1 тег w:trHeight будет удален
    /// </summary>
    public int Value {
      get {
        if (!XmlEl.HasAttribute("w:val"))
          return -1;
        return Int32.Parse(XmlEl.GetAttribute("w:val"));
      }
      set {
        if (value == -1)
          Delete();
        else
          XmlEl.SetAttribute("val", XmlEl.NamespaceURI, value.ToString());
      }
    }
  }

  public class TblBorders : Node {
    public TblBorders() : base("w:tblBorders") { }
    public TblBorders(Node parent) : base(parent, "w:tblBorders") { }
    public TblBorders(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:tblBorders") { }

    public void CompareBorders(Border Top, Border Bottom, Border Left, Border Right, Border InsideH, Border InsideV, string author = "TDV") {
      if (Top != this.Top || Bottom != this.Bottom || Left != this.Left || Right != this.Right || InsideH != this.InsideH || InsideV != this.InsideV) {
        ParentTable.CreateChangeNodes(author);
      }
      if (Top != this.Top)
        this.Top = Top;
      if (Bottom != this.Bottom)
        this.Bottom = Bottom;
      if (Left != this.Left)
        this.Left = Left;
      if (Right != this.Right)
        this.Right = Right;
      if (InsideV != this.InsideV)
        this.InsideV = InsideV;
      if (InsideH != this.InsideH)
        this.InsideH = InsideH;
    }
    private Table _parentTable;
    public Table ParentTable {
      get {
        if (_parentTable == null)
          _parentTable = GetParentRecurcieve<Table>();
        return _parentTable;
      }
    }

    public Border Left {
      get {
        return FindChild<Left>()?.Border ?? ParentTable?.TblPr?.Style?.FindChild<TableProp>()?.FindChild<TblBorders>()?.FindChild<Left>()?.Border;
      }
      set {
        FindChildOrCreate<Left>().Border = value;
      }
    }
    public Border Right {
      get {
        return FindChild<Right>()?.Border ?? ParentTable?.TblPr?.Style?.FindChild<TableProp>()?.FindChild<TblBorders>()?.FindChild<Right>()?.Border;
      }
      set { FindChildOrCreate<Right>().Border = value; }
    }
    public Border Top {
      get {
        return FindChild<Top>()?.Border ?? ParentTable?.TblPr?.Style?.FindChild<TableProp>()?.FindChild<TblBorders>()?.FindChild<Top>()?.Border;
      }
      set { FindChildOrCreate<Top>().Border = value; }
    }
    public Border Bottom {
      get {
        return FindChild<Bottom>()?.Border ?? ParentTable?.TblPr?.Style?.FindChild<TableProp>().FindChild<TblBorders>()?.FindChild<Bottom>()?.Border;
      }
      set { FindChildOrCreate<Bottom>().Border = value; }
    }
    public Border InsideH {
      get {
        return FindChild<InsideH>()?.Border ?? ParentTable?.TblPr?.Style?.FindChild<TableProp>()?.FindChild<TblBorders>()?.FindChild<InsideH>()?.Border;
      }
      set { FindChildOrCreate<InsideH>().Border = value; }
    }
    public Border InsideV {
      get {
        return FindChild<InsideV>()?.Border ?? ParentTable?.TblPr?.Style?.FindChild<TableProp>()?.FindChild<TblBorders>()?.FindChild<InsideV>()?.Border;
      }
      set { FindChildOrCreate<InsideV>().Border = value; }
    }
  }

  public class BorderNode : Node, IEquatable<BorderNode> {
    public BorderNode(string qualifiedName) : base(qualifiedName) { }
    public BorderNode(Node parent, string qualifiedName) : base(parent, qualifiedName) { }
    public BorderNode(XmlElement xmlElement, Node parent, string qualifiedName) : base(xmlElement, parent, qualifiedName) { }

    public Border Border {
      get {
        return new Border(LineType, Sz, Space, Color);
      }
      set {
        LineType = value.type;
        Sz = value.size;
        Space = value.space;
        Color = value.color;
      }
    }

    public bool Equals(BorderNode other) {
      if (this == null && other == null)
        return true;
      if ((other == null && this != null) || (other != null && this == null))
        return false;
      return Sz == other.Sz && Space == other.Space && LineType == other.LineType && Color == other.Color;
    }

    public LINE_TYPE LineType {
      get {
        if (HasAttribute("w:val"))
          return EnumExtentions.ToEnum<LINE_TYPE>(GetAttribute("w:val"));
        else
          return LINE_TYPE.NONE;
      }
      set {
        SetAttribute("w:val", value.ToStringValue());
      }
    }

    public int Sz {
      get {
        if (HasAttribute("w:sz"))
          return Int32.Parse(GetAttribute("w:sz"));
        return 0;
      }
      set {
        SetAttribute("w:sz", value.ToString());
        NodeChanded();
      }
    }

    public int Space {
      get {
        if (HasAttribute("w:space"))
          return Int32.Parse(GetAttribute("w:space"));
        return 0;
      }
      set {
        SetAttribute("w:space", value.ToString());
        NodeChanded();
      }
    }

    public string ThemeColor {
      get {
        if (HasAttribute("w:themeColor"))
          return GetAttribute("w:themeColor");
        else
          return null;
      }
    }

    public string Color {
      get {
        if (HasAttribute("w:color"))
          return GetAttribute("w:color");
        if (!string.IsNullOrEmpty(ThemeColor)) {
          DocxDocument docx = GetDocxDocument();
          return docx.ThemeDefault.ThemeElements.ClrScheme.ChildNodes.Where(x => x.XmlEl.Name == ThemeColor)
              .FirstOrDefault()?.FindChild<SrgbClr>()?.Value ?? "auto";
        }
        return "auto";
      }
      set {
        if (string.IsNullOrEmpty(value))
          RemoveAttribute("w:color");
        else
          SetAttribute("w:color", value.Replace("000000", "auto"));

        RemoveAttribute("w:themeColor");
        NodeChanded();
      }
    }
  }

  public class Top : BorderNode {
    public Top() : base("w:top") { }
    public Top(Node parent) : base(parent, "w:top") { }
    public Top(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:top") { }
  }
  public class Left : BorderNode {
    public Left() : base("w:left") { }
    public Left(Node parent) : base(parent, "w:left") { }
    public Left(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:left") { }
  }
  public class Bottom : BorderNode {
    public Bottom() : base("w:bottom") { }
    public Bottom(Node parent) : base(parent, "w:bottom") { }
    public Bottom(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:bottom") { }
  }
  public class Right : BorderNode {
    public Right() : base("w:right") { }
    public Right(Node parent) : base(parent, "w:right") { }
    public Right(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:right") { }
  }
  public class InsideH : BorderNode {
    public InsideH() : base("w:insideH") { }
    public InsideH(Node parent) : base(parent, "w:insideH") { }
    public InsideH(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:insideH") { }
  }
  public class InsideV : BorderNode {
    public InsideV() : base("w:insideV") { }
    public InsideV(Node parent) : base(parent, "w:insideV") { }
    public InsideV(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:insideV") { }
  }

  /// <summary>
  /// TableCell
  /// </summary>
  public class Tc : Node {
    private Table _paretnTable;
    public Tc() : base("w:tc") { }

    public override void InitXmlElement() {
      base.InitXmlElement();
      NewNodeFirst<TcProp>();
      NewNodeLast<Paragraph>();
    }

    //объединение колонок
    public int GridSpan {
      get {
        return TcProp.GridSpan.Value;
      }
      set {
        TcProp.GridSpan.Value = value;
      }
    }

    public Size Width {
      get {
        return TcProp.TcW.Width;
      }
      set {
        TcProp.TcW.Width = value;
      }
    }

    public Table ParentTable {
      get {
        if (_paretnTable == null)
          _paretnTable = GetParentRecurcieve<Table>();
        return _paretnTable;
      }
    }

    public Tc(Node parent, int cellIndex) : base(parent, "w:tc") { CellIndex = cellIndex; }
    public Tc(XmlElement xmlElement, Node parent, int cellIndex) : base(xmlElement, parent, "w:tc") { CellIndex = cellIndex; }
    public TcProp TcProp {
      get {
        return FindChildOrCreate<TcProp>();
      }
    }

    public void CorrectDel(string author = "TDV") {
      foreach (Paragraph p in Paragraphs)
        p.CorrectDel(author);
    }

    public int RowIndex {
      get {
        Tr row = GetParentRecurcieve<Tr>();
        return row.RowIndex;
      }
    }

    public readonly int CellIndex;

    public Tc GetLeftCell() {
      if (CellIndex - 1 < 0)
        return null;
      return ParentTable.GetCell(RowIndex, CellIndex - 1);
    }

    public Tc GetRightCell() {
      Tr row = ParentTable.Rows[RowIndex];
      if (row.Cells.Count() <= CellIndex + 1)
        return null;
      return ParentTable.GetCell(RowIndex, CellIndex + 1);
    }

    public Tc GetTopCell() {
      if (RowIndex == 0)
        return null;
      return ParentTable.GetCell(RowIndex - 1, CellIndex);
    }

    public Tc GetBottomCell() {
      if (ParentTable.Rows.Count() <= RowIndex + 1)
        return null;
      return ParentTable.GetCell(RowIndex + 1, CellIndex);
    }

    public List<Paragraph> Paragraphs {
      get { return ChildNodes.Where(x => x is Paragraph).Select(x => (Paragraph)x).ToList(); }
    }

    public override string Text {
      get {
        List<string> arr = new List<string>();
        foreach (Node n in ChildNodes) {
          if (n is Paragraph)
            arr.Add(((Paragraph)n).Text);
          else
            arr.Add(n.Text);
        }
        return string.Join(" ", arr);
      }
    }

    public override string ToString() {
      return Text;
    }

    public void CompareStyle(CellStyle style, string author = "TDV") {
      CompareWidth(style.width, author);
      CompareVAlign(style.vAlign, author);
      CompareBorders(style.borderLeft, style.borderRight, style.borderTop, style.borderBottom, true, author);
    }

    public void CompareWidth(Size width, string author = "TDV") {
      if (width != TcProp.TcW.Width) {
        ParentTable.CreateChangeNodes(author);
        TcProp.TcW.Width = width;
      }
    }

    public void CompareVAlign(VERTICAL_ALIGN vAlign, string author = "TDV") {
      if (vAlign != TcProp.vAlign.Align) {
        ParentTable.CreateChangeNodes(author);
        TcProp.vAlign.Align = vAlign;
      }
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="left"></param>
    /// <param name="right"></param>
    /// <param name="top"></param>
    /// <param name="bottom"></param>
    /// <param name="applyToNearCells">Применить стиль к смежным ячейкам</param>
    /// <param name="author"></param>
    public void CompareBorders(Border left = null, Border right = null, Border top = null, Border bottom = null, bool applyToNearCells = true, string author = "TDV") {
      if (left != null) {
        TcProp.TcBorders.CompareBorders(top, bottom, left, right, author);
        //TcProp.TcBorders.CompareBorder(BORDER_TYPE.LEFT, left, author);
        if (applyToNearCells) {
          Tc leftCell = GetLeftCell();
          if (leftCell != null)
            leftCell.CompareBorders(null, left, null, null, applyToNearCells: false, author);
        }
      }
      if (right != null) {
        TcProp.TcBorders.CompareBorders(top, bottom, left, right, author);
        if (applyToNearCells) {
          Tc rightCell = GetRightCell();
          if (rightCell != null)
            rightCell.CompareBorders(right, null, null, null, applyToNearCells: false, author);
        }
      }
      if (top != null) {
        TcProp.TcBorders.CompareBorders(top, bottom, left, right, author);
        if (applyToNearCells) {
          Tc topCell = GetTopCell();
          if (topCell != null)
            topCell.CompareBorders(null, null, null, top, applyToNearCells: false, author);
        }
      }
      if (bottom != null) {
        TcProp.TcBorders.CompareBorders(top, bottom, left, right, author);
        if (applyToNearCells) {
          Tc bottomCell = GetBottomCell();
          if (bottomCell != null)
            bottomCell.CompareBorders(null, null, bottom, null, applyToNearCells: false, author);
        }
      }
    }

    /// <summary>
    /// -1 - уже учавствует в другом мердже
    ///  0 - 
    /// </summary>
    public int RowSpan {
      get {
        return TcProp.vMerge.Value;
      }
      set {
        TcProp.vMerge.Value = value;
      }
    }

    /// <summary>
    ///  
    /// </summary>
    public int ColSpan {
      get {
        return TcProp.GridSpan.Value;
      }
      set {
        TcProp.GridSpan.Value = value;
      }
    }
  }

  public class TcW : Node {
    public TcW() : base("w:tcW") { }
    public TcW(Node parent) : base(parent, "w:tcW") { }
    public TcW(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:tcW") { }

    private Table _parentTable;
    public Table ParentTable {
      get {
        if (_parentTable == null)
          _parentTable = GetParentRecurcieve<Table>();
        return _parentTable;
      }
    }
    public Size Width {
      get {
        Size parentSize = null;
        if (SizeType == TABLE_WIDTH_TYPE.PCT)
          parentSize = ParentTable.Width;
        return new Size(Int32.Parse(GetAttribute("w:w")), SizeType, parentSize);
      }
      set {
        SizeType = value.SizeType;
        SetAttribute("w:w", value.Value.ToString());
      }
    }

    public TABLE_WIDTH_TYPE SizeType {
      get {
        TABLE_WIDTH_TYPE result = TABLE_WIDTH_TYPE.AUTO;
        Enum.TryParse<TABLE_WIDTH_TYPE>(XmlEl.GetAttribute("w:type"), true, out result);
        return result;
      }
      set { XmlEl.SetAttribute("type", XmlEl.NamespaceURI, value.ToString().ToLower()); }
    }
  }

  public class TcProp : Node {
    public TcProp() : base("w:tcPr") { }
    public TcProp(Node parent) : base(parent, "w:tcPr") { }
    public TcProp(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:tcPr") { }
    public TcW TcW {
      get {
        return FindChildOrCreate<TcW>();
      }
    }

    public TcBorders TcBorders {
      get {
        return FindChildOrCreate<TcBorders>();
      }
    }
    public VAlign vAlign {
      get {
        return FindChildOrCreate<VAlign>();
      }
    }
    public VMerge vMerge {
      get {
        return FindChildOrCreate<VMerge>();
      }
    }
    public GridSpan GridSpan {
      get {
        return FindChildOrCreate<GridSpan>();
      }
    }
  }

  public class VAlign : Node {
    public VAlign() : base("w:vAlign") { }
    public VAlign(Node parent) : base(parent, "w:vAlign") { }
    public VAlign(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:vAlign") { }

    /// <summary>
    /// При установке значения -1 тег w:trHeight будет удален
    /// </summary>
    public VERTICAL_ALIGN Align {
      get {
        VERTICAL_ALIGN result = VERTICAL_ALIGN.TOP;
        Enum.TryParse<VERTICAL_ALIGN>(XmlEl.GetAttribute("w:val"), true, out result);
        return result;
      }
      set {
        if (value == VERTICAL_ALIGN.TOP)
          Delete();
        else
          XmlEl.SetAttribute("val", XmlEl.NamespaceURI, value.ToString().ToLower());
      }
    }
  }

  public class VMerge : Node {
    public VMerge() : base("w:vMerge") { }
    public VMerge(Node parent) : base(parent, "w:vMerge") { }
    public VMerge(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:vMerge") { }

    /// <summary>
    /// -1 - пустой тег <w:vMerge/>. пустым тегом заполняются ячейки строками ниже
    /// </summary>
    public int Value {
      get {
        int result = -1;
        if (XmlEl.Attributes.Count == 0)
          return result;
        Int32.TryParse(XmlEl.GetAttribute("w:val"), out result);
        if (XmlEl.GetAttribute("w:val") == "restart")
          result = GetParentRecurcieve<Table>().Rows.Count;
        return result;
      }
      set {
        Table tbl = GetParentRecurcieve<Table>();
        int CurrRowIndex = GetParentRecurcieve<Tc>().RowIndex;
        int CurrColIndex = GetParentRecurcieve<Tc>().CellIndex;
        if (value == -1) {
          XmlEl.RemoveAllAttributes();
        }
        else if (value == 0 || value == 1) {
          int oldVal = Value;
          if (Value > 0)
            for (int rowIndex = CurrRowIndex + 1; rowIndex < Value; rowIndex++) {
              Tc cell = tbl.GetCell(rowIndex, CurrColIndex);
              cell.TcProp.vMerge.Delete();
            }
          //xmlEl.RemoveAllAttributes();
          Delete();
        }
        else {
          XmlEl.SetAttribute("val", XmlEl.NamespaceURI, value.ToString());

          ///установить для дстрок ниже тег <w:vMerge/>
          for (int rowIndex = CurrRowIndex + 1; rowIndex < CurrRowIndex + value; rowIndex++) {
            Tc cell = tbl.GetCell(rowIndex, CurrColIndex);
            cell.TcProp.vMerge.Value = -1;
          }
        }
      }
    }
  }

  public class GridSpan : Node {
    public GridSpan() : base("w:gridSpan") { }
    public GridSpan(Node parent) : base(parent, "w:gridSpan") { }
    public GridSpan(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:gridSpan") { }

    public int Value {
      get {
        if (HasAttribute("w:val"))
          try {
            return Int32.Parse(GetAttribute("w:val"));
          }
          catch { }
        return 1;
      }
      set {
        if (value == 1 || value == 0) {
          Delete();
        }
        else {
          SetAttribute("w:val", value.ToString());
        }
      }
    }
  }

  public class TcBorders : Node {
    public TcBorders() : base("w:tcBorders") { }
    public TcBorders(Node parent) : base(parent, "w:tcBorders") { }
    public TcBorders(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:tcBorders") { }

    private Table _parentTable;
    public Table ParentTable {
      get {
        if (_parentTable == null)
          _parentTable = GetParentRecurcieve<Table>();
        return _parentTable;
      }
    }

    public void CompareBorders(Border Top, Border Bottom, Border Left, Border Right, string author = "TDV") {
      if ((Top == null && Top != this.Top) || (Bottom != null && Bottom != this.Bottom) || (Left != null && Left != this.Left) || (Right != null && Right != this.Right)) {
        ParentTable.CreateChangeNodes(author);
      }
      if (Top != null && Top != this.Top)
        this.Top = Top;
      if (Bottom != null && Bottom != this.Bottom)
        this.Bottom = Bottom;
      if (Left != null && Left != this.Left)
        this.Left = Left;
      if (Right != null && Right != this.Right)
        this.Right = Right;
    }

    public Border Left {
      get {
        return FindChild<Left>()?.Border;// ?? ParentTable?.TblPr?.Style?.FindChild<TableProp>()?.FindChild<TblBorders>()?.FindChild<Left>()?.Border;
      }
      set {
        FindChildOrCreate<Left>().Border = value;
      }
    }
    public Border Right {
      get {
        return FindChild<Right>()?.Border;// ?? ParentTable?.TblPr?.Style?.FindChild<TableProp>()?.FindChild<TblBorders>()?.FindChild<Right>()?.Border;
      }
      set { FindChildOrCreate<Right>().Border = value; }
    }
    public Border Top {
      get {
        return FindChild<Top>()?.Border;// ?? ParentTable?.TblPr?.Style?.FindChild<TableProp>()?.FindChild<TblBorders>()?.FindChild<Top>()?.Border;
      }
      set { FindChildOrCreate<Top>().Border = value; }
    }
    public Border Bottom {
      get {
        return FindChild<Bottom>()?.Border;// ?? ParentTable?.TblPr?.Style?.FindChild<TableProp>().FindChild<TblBorders>()?.FindChild<Bottom>()?.Border;
      }
      set { FindChildOrCreate<Bottom>().Border = value; }
    }
  }
}