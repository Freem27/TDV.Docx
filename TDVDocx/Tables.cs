using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace TDV.Docx
{

    public class CellStyle
    {
        public CellStyle(
            VERTICAL_ALIGN vAlign,
                    Border borderLeft,
                    Border borderRight,
                    Border borderTop,
                    Border borderBottom,
                    int  width,
                    TABLE_WIDTH_TYPE widthType
            )
        {
            this.vAlign = vAlign;
            this.borderLeft = borderLeft;
            this.borderRight = borderRight;
            this.borderTop = borderTop;
            this.borderBottom = borderBottom;
            this.width = width;
            this.widthType = widthType;
        }
        public VERTICAL_ALIGN vAlign;
        public Border borderLeft;
        public Border borderRight;
        public Border borderTop;
        public Border borderBottom;
        public int width;
        public TABLE_WIDTH_TYPE widthType;
    }

    public class TableStyle
    {
        public TableStyle(
            Border borderLeft,
        Border borderRight,
        Border borderTop,
        Border borderBottom,
        Border borderInsideH,
        Border borderInsideV,
        bool applyBorderToCells,
        int width,
        TABLE_WIDTH_TYPE widthType,
        int indentingWidth,
        TABLE_WIDTH_TYPE indentingWidthType
        )
        {
            this.borderLeft        =borderLeft;
            this.borderRight       =borderRight;
            this.borderTop         =borderTop;
            this.borderBottom      =borderBottom;
            this.borderInsideH     =borderInsideH;
            this.borderInsideV     =borderInsideV;
            this.applyBorderToCells = applyBorderToCells;
            this.width             =width;
            this.widthType         =widthType;
            this.indentingWidth    =indentingWidth;
            this.indentingWidthType=indentingWidthType;
        }

        public Border borderLeft;
        public Border borderRight;
        public Border borderTop;
        public Border borderBottom;
        public Border borderInsideH;
        public Border borderInsideV;
        public bool applyBorderToCells;
        public int width;
        public TABLE_WIDTH_TYPE widthType;
        public int indentingWidth;
        public TABLE_WIDTH_TYPE indentingWidthType;
    }
    public class Table : Node
    {
        public Table() : base("w:tbl") { }
        public Table(Node parent) : base(parent, "w:tbl") { }
        public Table(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:tbl") { }

        /// <summary>
        /// Создает ChangeNode  для tblPr и  tblGrid
        /// </summary>
        internal void CreateChangeNodes(string author = "TDV")
        {
            tblPr.CreateChangeNode("w:tblPrChange", tblPr.xmlEl, author);
            tblGrid.CreateChangeNode("w:tblGridChange", tblGrid.xmlEl, author);
            foreach (Tr row in Rows)
            {
                row.trPr.CreateChangeNode(author);
                foreach (Tc cell in row.Cells)
                    cell.tcProp.CreateChangeNode(author);
            }
        }
        public void ApplyAllFixes()
        {
            
            FindChild<TableProp>()?.FindChild<TblPrChange>()?.Delete();
            FindChild<TableGrid>()?.FindChild<TblGridChange>()?.Delete();

            foreach (Tr tr in Rows)
            {
                if (tr.FindChild<TrProp>()?.FindChild<Del>() != null)
                {
                    tr.Delete();
                    continue;
                }
                tr.FindChild<TrProp>()?.FindChild<Ins>()?.Delete();
                tr.FindChild<TrProp>()?.FindChild<TrPrChange>()?.Delete();

                foreach (Tc tc in tr.Cells)
                {
                    tc.FindChild<TcProp>()?.FindChild<TcPrChange>()?.Delete();
                    foreach(Node n in tc.childNodes)
                    {
                        if (n is Paragraph)
                            ((Paragraph)n).ApplyAllFixes();
                        else if (n is Table)
                            ((Table)n).ApplyAllFixes();
                    }
                }
            }
        }
        public void CompateStyle(TableStyle style, string author = "TDV")
        {
            CompareBorders(style.borderLeft, style.borderRight, style.borderTop, style.borderBottom,
                style.borderInsideH, style.borderInsideV, style.applyBorderToCells, author);
            CompareIndenting(style.indentingWidth, style.indentingWidthType, author);
            CompareTblWidth(style.width, style.widthType, author);
        }
        public TableProp tblPr
        {
            get
            {
                TableProp result = childNodes.Where(x => x is TableProp).Select(x => (TableProp)x).FirstOrDefault();
                if (result == null)
                    result = new TableProp(this);
                return result;
            }
        }

        public void CompareBorders(Border left, Border right, Border top, Border bottom, Border insideH, Border insideV, bool applyToCells, string author = "TDV")
        {
            tblPr.tblBorders.CompareBorder(BORDER.LEFT, left, author);
            tblPr.tblBorders.CompareBorder(BORDER.RIGHT, right, author);
            tblPr.tblBorders.CompareBorder(BORDER.TOP, top, author);
            tblPr.tblBorders.CompareBorder(BORDER.BOTTOM, bottom, author);
            tblPr.tblBorders.CompareBorder(BORDER.INSIDE_H, insideH, author);
            tblPr.tblBorders.CompareBorder(BORDER.INSIDE_V, insideV, author);
            foreach(Tr tr in Rows)
            foreach (Tc tc in tr.Cells)
            {
                if (applyToCells)
                    tc.CompareBorders(left, right, top, bottom, false, author);
                tc.tcProp.CreateChangeNode(author);
            }
        }

        public void CompareIndenting(int width, TABLE_WIDTH_TYPE type,string author)
        {
            tblPr.CompareIndenting(width, type, author);
        }

        public void CompareTblWidth(int width, TABLE_WIDTH_TYPE type,string author)
        {
            tblPr.CompareTblWidth(width, type, author);
        }
        public TableGrid tblGrid
        {
            get
            {
                var result = childNodes.Where(x => x is TableGrid).Select(x => x).FirstOrDefault();
                if (result == null)
                    result = new TableGrid(this);
                return (TableGrid)result;
            }
        }

        public Tc GetCell(int row, int col)
        {
            if (Rows.Count - 1<row)
                return null;
            if (Rows[row].Cells.Count - 1 < col)
                return null;
            return Rows[row].Cells[col];
        }

        public List<Tr> Rows
        {
            get { return childNodes.Where(x => x is Tr).Select(x => (Tr) x).ToList(); }
        }


        public override string ToString()
        {
            string result = "";
            foreach (Tr row in Rows)
            foreach (Tc cell in row.Cells)
            foreach (Paragraph p in cell.Paragraphs)
                result += p.ToString();
            return result;
        }
    }

    public class TableProp : Node
    {
        public TableProp() : base("w:tblPr") { }
        public TableProp(Node parent) : base(parent, "w:tblPr") { }
        public TableProp(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:tblPr") { }

        public TblBorders tblBorders
        {
            get
            {
                var result = childNodes.Where(x => x is TblBorders).Select(x => x).FirstOrDefault();
                if (result == null)
                    result = new TblBorders(this);
                return (TblBorders)result;
            }
        }

        public Style CurrStyle {
            get
            {
                TblStyle el = (TblStyle) childNodes.Where(x=>x is TblStyle).FirstOrDefault();
                if (el == null)
                    return null;
                return GetDocxDocument().styles.GetStyleById(el.StyleId);
            }
        }

        public void CompareIndenting(int width, TABLE_WIDTH_TYPE type,string author)
        {
            if (tblInd.Width != width || tblInd.type != type)
            {
                Table tbl = GetParentRecurcieve<Table>();
                tbl.CreateChangeNodes(author);
                tblInd.Width=width;
                tblInd.type = type;
            }
        }

        public void CompareTblWidth(int width, TABLE_WIDTH_TYPE type, string author)
        {
            if (tblW.Width != width || tblW.type != type)
            {
                Table tbl = GetParentRecurcieve<Table>();
                tbl.CreateChangeNodes(author);
                tblW.Width = width;
                tblW.type = type;
            }
        }

        public TableWidth tblW
        {
            get
            {
                var result = childNodes.Where(x => x is TableWidth).Select(x => x).FirstOrDefault();
                if (result == null)
                    result = new TableWidth(this);
                return (TableWidth)result;
            }
        }

        public TblInd tblInd
        {
            get
            {
                var result = childNodes.Where(x => x is TblInd).Select(x => x).FirstOrDefault();
                if (result == null)
                    result = new TblInd(this);
                return (TblInd)result;
            }
        }

        public TblStyle tblStyle
        {
            get
            {
                var result = childNodes.Where(x => x is TblStyle).Select(x => x).FirstOrDefault();
                if (result == null)
                    result = new TblStyle(this);
                return (TblStyle)result;
            }
        }

    }

    public class TblStyle : Node
    {
        public TblStyle() : base("w:tblStyle") { }
        public TblStyle(Node parent) : base(parent, "w:tblStyle") { }
        public TblStyle(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:tblStyle") { }

        public string StyleId
        {
            get { return xmlEl.GetAttribute("w:val"); }
        }

    }
    public enum TABLE_WIDTH_TYPE
    {
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
        /// Значение в процентах от ширины таблицы 
        /// </summary>
        PTC
    }
    public class TableWidth : Node
    {
        public TableWidth() : base("w:tblW") { }
        public TableWidth(Node parent) : base(parent, "w:tblW") { }
        public TableWidth(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:tblW") { }

        public int Width
        {
            get
            {
                return Int32.Parse(xmlEl.GetAttribute("w:w"));
            }
            set { xmlEl.SetAttribute("w", xmlEl.NamespaceURI, value.ToString()); }
        }

        public TABLE_WIDTH_TYPE type
        {
            get
            {
                return (TABLE_WIDTH_TYPE)Enum.Parse(typeof(TABLE_WIDTH_TYPE), xmlEl.GetAttribute("w:type"),true);
            }
            set { xmlEl.SetAttribute("type", xmlEl.NamespaceURI, value.ToString().ToLower()); }
        }

    }

    /// <summary>
    /// Отступ таблицы (слева)
    /// </summary>
    public class TblInd : Node
    {
        public TblInd() : base("w:tblInd") { }
        public TblInd(Node parent) : base(parent, "w:tblInd") { }
        public TblInd(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:tblInd") { }

        public int Width
        {
            get
            {
                if (!xmlEl.HasAttribute("w:w"))
                    return 0;
                return Int32.Parse(xmlEl.GetAttribute("w:w"));
            }
            set { xmlEl.SetAttribute("w", xmlEl.NamespaceURI, value.ToString()); }
        }

        public TABLE_WIDTH_TYPE type
        {
            get
            {
                if (!xmlEl.HasAttribute("w:type"))
                    return TABLE_WIDTH_TYPE.NIL;
                return (TABLE_WIDTH_TYPE)Enum.Parse(typeof(TABLE_WIDTH_TYPE), xmlEl.GetAttribute("w:type"), true);
            }
            set { xmlEl.SetAttribute("type", xmlEl.NamespaceURI, value.ToString().ToLower()); }
        }

    }

    public class TableGrid : Node
    {
        public TableGrid() : base("w:tblGrid") { }
        public TableGrid(Node parent) : base(parent, "w:tblGrid") { }
        public TableGrid(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:tblGrid") { }

        public List<GridColumn> gridCols
        {
            get { return childNodes.Where(x => x is GridColumn).Select(x => (GridColumn) x).ToList(); }
        }
    }

    public class GridColumn : Node
    {
        public GridColumn() : base("w:gridCol") { }
        public GridColumn(Node parent) : base(parent, "w:gridCol") { }
        public GridColumn(XmlElement xmlElement, Node parent, int colIndex) : base(xmlElement, parent, "w:gridCol"){ColIndex = colIndex; }
        public int Width
        {
            get
            {
                return Int32.Parse(xmlEl.GetAttribute("w:w"));
            }
            set
            {
                xmlEl.SetAttribute("w", xmlEl.NamespaceURI, value.ToString());
            }
        }

        public readonly int ColIndex;
        public TABLE_WIDTH_TYPE type
        {
            get
            {
                TABLE_WIDTH_TYPE result = TABLE_WIDTH_TYPE.AUTO;
                Enum.TryParse<TABLE_WIDTH_TYPE>(xmlEl.GetAttribute("w:type"), true,out result);
                return result;
            }
            set { xmlEl.SetAttribute("type", xmlEl.NamespaceURI, value.ToString().ToLower()); }
        }
        /// <summary>
        /// Задает значения ширины таблицы в режиме правки
        /// </summary>
        /// <param name="width">Значение в нормальных процентах (от 0 до 100). будет переведено в диапазон от 0 до 5000</param>
        /// <param name="type"></param>
        /// <param name="applyToColumns">Применить к столбцам таблицы</param>
        /// <param name="author"></param>
        public void CompareWidth(int width, TABLE_WIDTH_TYPE type,bool applyToColumns,string author="TDV")
        {
            if (width != Width || type!=this.type)
            {
                Table table = GetParentRecurcieve<Table>();
                table.CreateChangeNodes(author);
                if (type == TABLE_WIDTH_TYPE.PTC)
                    width = width * 50;
                Width = width;
                this.type = type;
                if(applyToColumns)
                    foreach (Tr row in table.Rows)
                        row.Cells[this.ColIndex].CompareWidth(100, TABLE_WIDTH_TYPE.PTC, author);
            }
        }
    }

    public class Tr : Node
    {
        public Tr() : base("w:tr") { }
        public Tr(Node parent, int rowIndex) : base(parent, "w:tr") { RowIndex = rowIndex; }

        public Tr(XmlElement xmlElement, Node parent, int rowIndex) : base(xmlElement, parent, "w:tr") { RowIndex = rowIndex; }

        public TrProp trPr {
            get
            {
                var result = childNodes.Where(x => x is TrProp).Select(x => x).FirstOrDefault();
                if (result == null)
                    result = new TrProp(this);
                return (TrProp)result;
            }
        }

        public readonly int RowIndex;

        public List<Tc> Cells
        {
            get { return childNodes.Where(x => x is Tc).Select(x => (Tc)x).ToList(); }
        }

        public void CorrectDel(string author = "TDV")
        {
            trPr.SetCorrectionMode("del", author);
            foreach(Tc cell in Cells)
                cell.CorrectDel(author);
        }


        public void CompareHeigth(int heigth, string author="TDV")
        {
            trPr.CompareHeigth(heigth, author);
        }

        public override string ToString()
        {
            return Text;
        }
    }

    public class TrProp : Node,ICorrectable
    {
        public TrProp() : base("w:trPr") { }
        public TrProp(Node parent) : base(parent, "w:trPr") { }
        public TrProp(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:trPr") { }



        public void CreateChangeNode(string author="TDV")
        {
            CreateChangeNode("w:trPrChange", xmlEl, author);
        }
        /// <summary>
        /// Изменение высоты строки в режиме правки
        /// </summary>
        /// <param name="heigth">-1: тег будет удален</param>
        /// <param name="author"></param>
        public void CompareHeigth(int heigth,string author)
        {
            if ( trHeight.Height != heigth)
            {
                CreateChangeNode(author);
                trHeight.Height = heigth;
            }
        }


        public TrHeight trHeight
        {
            get
            {
                var result = childNodes.Where(x => x is TrHeight).Select(x => x).FirstOrDefault();
                if (result == null)
                    result = new TrHeight(this);
                return (TrHeight)result;
            }
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
                cNode = (XmlElement)doc.CreateElement($"w:{mode}", xmlEl.NamespaceURI);
                cNode.SetAttribute("id", xmlEl.NamespaceURI, (doc.GetLastId() + 1).ToString());
                xmlEl.InsertBefore(cNode, xmlEl.FirstChild);
            }
            cNode.SetAttribute("author", xmlEl.NamespaceURI, author);
            cNode.SetAttribute("date", xmlEl.NamespaceURI, DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ssZ"));
        }

    }

    public class TrHeight : Node
    {
        public TrHeight() : base("w:trHeight") { }
        public TrHeight(Node parent) : base(parent, "w:trHeight") { }
        public TrHeight(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:trHeight") { }

        /// <summary>
        /// При установке значения -1 тег w:trHeight будет удален
        /// </summary>
        public int Height
        {
            get
            {
                if (!xmlEl.HasAttribute("w:val"))
                    return -1;
                return Int32.Parse(xmlEl.GetAttribute("w:val"));
            }
            set
            {
                if (value == -1)
                    Delete();
                else
                    xmlEl.SetAttribute("val", xmlEl.NamespaceURI, value.ToString());
            }
        }
    }
    
    public class TblBorders : Node
    {
        public TblBorders() : base("w:tblBorders") { }
        public TblBorders(Node parent) : base(parent, "w:tblBorders") { }
        public TblBorders(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:tblBorders") { }

        public void CompareBorder(BORDER type, Border b, string author = "TDV")
        {
            Border currBorder = GetBorder(type);

            if (currBorder != b)
            {
                Table table = GetParentRecurcieve<Table>();
                table.CreateChangeNodes(author);
                //parent.CreateChangeNode("w:tblPrChange", (XmlElement)xmlEl.ParentNode, author);
                //table.tblGrid.CreateChangeNode("w:tblGridChange", table.tblGrid.xmlEl, author);
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
        public Border InsideH
        {
            get { return GetBorder(BORDER.INSIDE_H); }
            set { SetBorder(BORDER.INSIDE_H, value); }
        }
        public Border InsideV
        {
            get { return GetBorder(BORDER.INSIDE_V); }
            set { SetBorder(BORDER.INSIDE_V, value); }
        }

        private Border GetBorder(BORDER type)
        {

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
                case BORDER.INSIDE_H:
                    localName = "insideH";
                    break;
                case BORDER.INSIDE_V:
                    localName = "insideV";
                    break;
                default:
                    break;
            }
            Style style = ((TableProp)parent).CurrStyle;
            XmlElement n = (XmlElement)xmlEl.SelectSingleNode($"w:{localName}", nsmgr);
            Border b = new Border();
            if (n == null && style != null)
            {
                TableProp styleTableProp = style.GetStyleProp<TableProp>();
                if(styleTableProp != null)
                    n= (XmlElement)styleTableProp.xmlEl.SelectSingleNode($"w:tblBorders/w:{localName}", nsmgr);
            }
            if (n != null)
            {
                b.color = n.GetAttribute("w:color");
                b.size = Int32.Parse(n.GetAttribute("w:sz"));
                if (n.HasAttribute("w:space"))
                    b.space = Int32.Parse(n.GetAttribute("w:space"));
                else b.space = -1;

                b.type = LINE_TYPE.UNKNOWN;
                Enum.TryParse<LINE_TYPE>(n.GetAttribute("w:val"), true, out b.type);
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
                case BORDER.INSIDE_H:
                    localName = "insideH";
                    break;
                case BORDER.INSIDE_V:
                    localName = "insideV";
                    break;
                default:
                    break;
            }
            XmlElement n = (XmlElement)xmlEl.SelectSingleNode($"{prefix}:{localName}", nsmgr);
            
            if (n == null)
            {
                n = doc.CreateElement($"{prefix}:{localName}", doc.DocumentElement.NamespaceURI);
                xmlEl.AppendChild(n);
            }

            if (b.type != LINE_TYPE.UNKNOWN)
                n.SetAttribute("val", xmlEl.NamespaceURI, b.type.ToString().ToLower());
            n.SetAttribute("sz", xmlEl.NamespaceURI, b.size.ToString());
            n.SetAttribute("space", xmlEl.NamespaceURI, b.space.ToString());
            n.SetAttribute("color", xmlEl.NamespaceURI, b.color);
        }
    }
    
    /// <summary>
    /// TableCell
    /// </summary>
    public class Tc : Node
    {
        public Tc() : base("w:tc")
        {
        }

        public override void InitXmlElement()
        {
            base.InitXmlElement();
            NewNodeFirst<TcProp>();
            NewNodeLast<Paragraph>();
        }

        public Tc(Node parent, int cellIndex) : base(parent, "w:tc") { CellIndex = cellIndex; }
        public Tc(XmlElement xmlElement, Node parent, int cellIndex) : base(xmlElement, parent, "w:tc") { CellIndex = cellIndex; }
        public TcProp tcProp
        {
            get
            {
                var result = childNodes.Where(x => x is TcProp).Select(x => x).FirstOrDefault();
                if (result == null)
                    result = new TcProp(this);
                return (TcProp)result;
            }
        }
        public void CorrectDel(string author = "TDV")
        {
            foreach (Paragraph p in Paragraphs)
                p.CorrectDel(author);
        }

        public int RowIndex
        {
            get { 
                Tr row=GetParentRecurcieve<Tr>();
                return row.RowIndex;
            }
        }

        public readonly int CellIndex;

        public Tc GetLeftCell()
        {
            if (CellIndex-1 < 0)
                return null;
            Table table = GetParentRecurcieve<Table>();
            return table.GetCell(RowIndex, CellIndex-1);
        }

        public Tc GetRightCell()
        {
            Table table = GetParentRecurcieve<Table>();
            Tr row = table.Rows[RowIndex];
            if (row.Cells.Count() <= CellIndex + 1)
                return null;
            return table.GetCell(RowIndex, CellIndex + 1);
        }

        public Tc GetTopCell()
        {
            if (RowIndex == 0)
                return null;
            Table table = GetParentRecurcieve<Table>();
            return table.GetCell(RowIndex - 1, CellIndex);
        }

        public Tc GetBottomCell()
        {
            Table table = GetParentRecurcieve<Table>();
            if (table.Rows.Count() <= RowIndex + 1)
                return null;
            return table.GetCell(RowIndex + 1, CellIndex);
        }

        public List<Paragraph> Paragraphs
        {
            get { return childNodes.Where(x => x is Paragraph).Select(x => (Paragraph) x).ToList(); }
        }

        public override string Text
        {
            get
            {
                List<string> arr = new List<string>();
                foreach (Node n in childNodes)
                {
                    if (n is Paragraph)
                        arr.Add(((Paragraph) n).Text);
                    else
                        arr.Add(n.Text);
                }
                return string.Join(" ", arr);
            }
        }


        public override string ToString()
        {
            return Text;
        }

        public void CompareStyle(CellStyle style, string author = "TDV")
        {
            CompareWidth(style.width, style.widthType, author);
            CompareVAlign(style.vAlign, author);
            CompareBorders(style.borderLeft, style.borderRight, style.borderTop, style.borderBottom, true, author);
        }

        public void CompareWidth(int width, TABLE_WIDTH_TYPE type, string author = "TDV")
        {
            if (width != tcProp.tcW.Width || type != tcProp.tcW.type)
            {
                Table table = GetParentRecurcieve<Table>();
                table.CreateChangeNodes(author);
                tcProp.CreateChangeNode("w:tcPrChange", tcProp.xmlEl, author);
                if (type == TABLE_WIDTH_TYPE.PTC)
                    width = width * 50;
                tcProp.tcW.Width = width;
                tcProp.tcW.type = type;
            }
        }

        public void CompareVAlign(VERTICAL_ALIGN vAlign, string author = "TDV")
        {
            if (vAlign != tcProp.vAlign.Align )
            {
                Table table = GetParentRecurcieve<Table>();
                table.CreateChangeNodes(author);
                tcProp.CreateChangeNode("w:tcPrChange", tcProp.xmlEl, author);
                tcProp.vAlign.Align = vAlign;
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
        public void CompareBorders(Border left=null, Border right = null, Border top = null, Border bottom = null,bool applyToNearCells=true, string author = "TDV")
        {
            if (left != null)
            {
                tcProp.tcBorders.CompareBorder(BORDER.LEFT, left, author);
                if(applyToNearCells)
                { 
                    Tc leftCell = GetLeftCell();
                    if (leftCell != null)
                        leftCell.CompareBorders(null, left,null,null, applyToNearCells:false, author);
                }
            }
            if(right!=null)
            { 
                tcProp.tcBorders.CompareBorder(BORDER.RIGHT, right, author); 
                if (applyToNearCells)
                {
                    Tc rightCell = GetRightCell();
                    if (rightCell != null)
                        rightCell.CompareBorders(right, null, null, null, applyToNearCells: false, author);
                }
            }
            if(top!=null)
            { 
                tcProp.tcBorders.CompareBorder(BORDER.TOP, top, author);
                if (applyToNearCells)
                {
                    Tc topCell = GetTopCell();
                    if (topCell != null)
                        topCell.CompareBorders(null, null, null, top, applyToNearCells: false, author);
                }
            }
            if(bottom!=null)
            { 
                tcProp.tcBorders.CompareBorder(BORDER.BOTTOM, bottom, author);
                if (applyToNearCells)
                {
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
        public int RowSpan
        {
            get
            {
                return tcProp.vMerge.val;
            }
            set
            {
                tcProp.vMerge.val = value;
            }
        }
        /// <summary>
        ///  
        /// </summary>
        public int ColSpan
        {
            get
            {
                return tcProp.gridSpan.val;
            }
            set
            {
                tcProp.gridSpan.val = value;
            }
        }
    }

    public class TcW : Node
    {
        public TcW() : base("w:tcW") { }
        public TcW(Node parent) : base(parent, "w:tcW") { }
        public TcW(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:tcW") { }

        public int Width
        {
            get
            {
                if (!xmlEl.HasAttribute("w:w"))
                    return -1;
                return Int32.Parse(xmlEl.GetAttribute("w:w"));
            }
            set
            {
                xmlEl.SetAttribute("w", xmlEl.NamespaceURI, value.ToString());
            }
        }

        public TABLE_WIDTH_TYPE type
        {
            get
            {
                TABLE_WIDTH_TYPE result = TABLE_WIDTH_TYPE.AUTO;
                Enum.TryParse<TABLE_WIDTH_TYPE>(xmlEl.GetAttribute("w:type"), true, out result);
                return result;
            }
            set { xmlEl.SetAttribute("type", xmlEl.NamespaceURI, value.ToString().ToLower()); }
        }
    }

    public class TcProp : Node
    {
        public TcProp() : base("w:tcPr") { }
        public TcProp(Node parent) : base(parent, "w:tcPr") { }
        public TcProp(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:tcPr") { }
        public TcW tcW
        {
            get
            {
                var result = childNodes.Where(x => x is TcW).Select(x => x).FirstOrDefault();
                if (result == null)
                    result = new TcW(this);
                return (TcW)result;
            }
        }

        public void CreateChangeNode(string author)
        { 
            CreateChangeNode("w:tcPrChange",  xmlEl, author);
        }

        public TcBorders tcBorders
        {
            get
            {
                TcBorders result = childNodes.Where(x => x is TcBorders).Select(x => (TcBorders)x).FirstOrDefault();
                if(result == null)
                    result = new TcBorders(this);
                return result;
            }
        }
        public VAlign vAlign
        {
            get
            {
                VAlign result = childNodes.Where(x => x is VAlign).Select(x => (VAlign)x).FirstOrDefault();
                if (result == null)
                    result = new VAlign(this);
                return result;
            }
        }
        public VMerge vMerge
        {
            get
            {
                VMerge result = childNodes.Where(x => x is VMerge).Select(x => (VMerge)x).FirstOrDefault();
                if (result == null)
                    result = new VMerge(this);
                return result;
            }
        }
        public GridSpan gridSpan
        {
            get
            {
                GridSpan result = childNodes.Where(x => x is GridSpan).Select(x => (GridSpan)x).FirstOrDefault();
                if (result == null)
                    result = new GridSpan(this);
                return result;
            }
        }
    }

    public enum VERTICAL_ALIGN { TOP,CENTER,BOTTOM}
    public class VAlign : Node
    {
        public VAlign() : base("w:vAlign") { }
        public VAlign(Node parent) : base(parent, "w:vAlign") { }
        public VAlign(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:vAlign") { }

        /// <summary>
        /// При установке значения -1 тег w:trHeight будет удален
        /// </summary>
        public VERTICAL_ALIGN Align
        {
            get
            {
                VERTICAL_ALIGN result = VERTICAL_ALIGN.TOP;
                Enum.TryParse<VERTICAL_ALIGN>(xmlEl.GetAttribute("w:val"),true,out result);
                return result;
            }
            set
            {
                if (value == VERTICAL_ALIGN.TOP)
                    Delete();
                else
                    xmlEl.SetAttribute("val", xmlEl.NamespaceURI, value.ToString().ToLower());
            }
        }
    }

    public class VMerge : Node
    {
        public VMerge() : base("w:vMerge") { }
        public VMerge(Node parent) : base(parent, "w:vMerge") { }
        public VMerge(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:vMerge") { }

        /// <summary>
        /// -1 - пустой тег <w:vMerge/>. пустым тегом заполняются ячейки строками ниже
        /// </summary>
        public int val
        {
            get
            {
                int result = -1;
                if (xmlEl.Attributes.Count == 0)
                    return result;
                Int32.TryParse(xmlEl.GetAttribute("w:val"), out result);
                if (xmlEl.GetAttribute("w:val") == "restart")
                    result = GetParentRecurcieve<Table>().Rows.Count;
                return result;
            }
            set
            {
                Table tbl = GetParentRecurcieve<Table>();
                int CurrRowIndex = GetParentRecurcieve<Tc>().RowIndex;
                int CurrColIndex = GetParentRecurcieve<Tc>().CellIndex;
                if (value == -1)
                {
                    xmlEl.RemoveAllAttributes();
                }
                else if (value == 0 || value == 1)
                {
                    int oldVal = val;
                    if (val > 0)
                        for (int rowIndex = CurrRowIndex + 1; rowIndex < val; rowIndex++)
                        {
                            Tc cell = tbl.GetCell(rowIndex, CurrColIndex);
                            cell.tcProp.vMerge.Delete();
                        }
                    //xmlEl.RemoveAllAttributes();
                    Delete();
                }
                else
                {
                    xmlEl.SetAttribute("val", xmlEl.NamespaceURI, value.ToString());

                    ///установить для дстрок ниже тег <w:vMerge/>
                    for (int rowIndex = CurrRowIndex + 1; rowIndex < CurrRowIndex + value; rowIndex++)
                    {
                        Tc cell = tbl.GetCell(rowIndex, CurrColIndex);
                        cell.tcProp.vMerge.val = -1;
                    }
                }
            }
        }
    }

    public class GridSpan : Node
    {
        public GridSpan() : base("w:gridSpan") { }
        public GridSpan(Node parent) : base(parent, "w:gridSpan") { }
        public GridSpan(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:gridSpan") { }

        /// <summary>
        /// 
        /// </summary>
        public int val
        {
            get
            {
                int result = 0;
                if (xmlEl.Attributes.Count == 0)
                    return result;
                Int32.TryParse(xmlEl.GetAttribute("w:val"), out result);
                //if (xmlEl.GetAttribute("w:val") == "restart")
                //    result = GetParentRecurcieve<Table>().Rows.Count;
                return result;
            }
            set
            {
                Table tbl = GetParentRecurcieve<Table>();
                int CurrRowIndex = GetParentRecurcieve<Tc>().RowIndex;
                int CurrColIndex = GetParentRecurcieve<Tc>().CellIndex;
                if (value == -1)
                {
                    xmlEl.RemoveAllAttributes();
                }
                else if (value == 0 || value == 1)
                {
                    if (val > 1)
                    {
                        Tc currCell= GetParentRecurcieve<Tc>();
                        Tr currRow = GetParentRecurcieve<Tr>();
                        int newRowWidth = currCell.tcProp.tcW.Width / val;
                        currCell.tcProp.gridSpan.Delete();
                        for (int cellIndex = CurrColIndex + 1; cellIndex < CurrColIndex + val; cellIndex++)
                        {
                            Tc newCell=currRow.NewNodeAfter<Tc>(currCell.xmlEl);
                            //newCell.tcProp.tcW.Width = newRowWidth;
                            //newCell.tcProp.tcW.type = currCell.tcProp.tcW.type;
                        }
                    }

                    Delete();
                }
                else
                {
                    xmlEl.SetAttribute("val", xmlEl.NamespaceURI, value.ToString());
                    Tr row = GetParentRecurcieve<Tr>();
                    ///установить для строк ниже тег <w:vMerge/>
                    for (int cellIndex = CurrColIndex + 1; cellIndex < CurrColIndex + value; cellIndex++)
                    {
                        Tc cell = tbl.GetCell(CurrRowIndex, cellIndex);
                        cell.Delete();
                    }
                }
            }
        }
    }

    public class TcBorders : Node
    {
        public TcBorders() : base("w:tcBorders") { }
        public TcBorders(Node parent) : base(parent, "w:tcBorders") { }
        public TcBorders(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:tcBorders") { }

        public void CompareBorder(BORDER type, Border b, string author = "TDV")
        {
            Border currBorder = GetBorder(type);

            if (currBorder != b)
            {
                Table table= GetParentRecurcieve<Table>();
                
                table.CreateChangeNodes(author);
                //table.tblPr.CreateChangeNode("w:tblPrChange", table.tblPr.xmlEl, author);
                //table.tblGrid.CreateChangeNode("w:tblGridChange", table.tblGrid.xmlEl, author);
                
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


        private Border GetBorder(BORDER type)
        {

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
                default:
                    break;
            }
            Style style = (GetParentRecurcieve<Table>().tblPr).CurrStyle;
            XmlElement n = (XmlElement)xmlEl.SelectSingleNode($"w:{localName}", nsmgr);
            Border b = new Border();
            if (n == null && style != null)
            {
                TableProp styleTableProp = style.GetStyleProp<TableProp>();
                if (styleTableProp != null)
                    n = (XmlElement)styleTableProp.xmlEl.SelectSingleNode($"w:tblBorders/w:{localName}", nsmgr);
            }
            if (n != null)
            {
                if (n.HasAttribute("w:color"))
                    b.color = n.GetAttribute("w:color");
                if (n.HasAttribute("color"))
                    b.color = n.GetAttribute("color");
                if (n.HasAttribute("sz"))
                    b.size = Int32.Parse(n.GetAttribute("sz"));
                if (n.HasAttribute("w:sz"))
                    b.size = Int32.Parse(n.GetAttribute("w:sz"));
                if (n.HasAttribute("w:space"))
                    b.space = Int32.Parse(n.GetAttribute("w:space"));
                if (n.HasAttribute("space"))
                    b.space = Int32.Parse(n.GetAttribute("space"));
                else b.space = -1;
                b.type = LINE_TYPE.UNKNOWN;
                if (n.HasAttribute("w:val"))
                    Enum.TryParse<LINE_TYPE>(n.GetAttribute("w:val"), true, out b.type);
                if (n.HasAttribute("val"))
                    Enum.TryParse<LINE_TYPE>(n.GetAttribute("val"), true, out b.type);
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
                case BORDER.INSIDE_H:
                    localName = "insideH";
                    break;
                case BORDER.INSIDE_V:
                    localName = "insideV";
                    break;
                default:
                    break;
            }
            XmlElement n = (XmlElement)xmlEl.SelectSingleNode($"{prefix}:{localName}", nsmgr);

            if (n == null)
            {
                n = doc.CreateElement($"{prefix}:{localName}", doc.DocumentElement.NamespaceURI);
                xmlEl.AppendChild(n);
            }

            if (b.type == LINE_TYPE.NONE)
            {
                n.SetAttribute("val", xmlEl.NamespaceURI, "nil");
            }
            else if (b.type != LINE_TYPE.UNKNOWN)
                n.SetAttribute("val", xmlEl.NamespaceURI, b.type.ToString().ToLower());

            n.SetAttribute("sz", xmlEl.NamespaceURI, b.size.ToString());
            n.SetAttribute("space", xmlEl.NamespaceURI, b.space.ToString());
            n.SetAttribute("color", xmlEl.NamespaceURI, b.color);
        }
    }

}
