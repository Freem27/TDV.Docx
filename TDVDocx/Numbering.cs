using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace TDV.Docx
{
    public class AbstractNumNotFoundException : Exception
    {
        public AbstractNumNotFoundException(string message) : base(message) { }
    }
    public class Numbering:BaseNode
    {
        public Numbering(DocxDocument docx):base(docx)
        {
            try
            {
                file = docx.sourceFolder.FindFile("numbering.xml", @"word");
                List<Section> sections = DocxDocument.Document.Sections;
                XmlDoc = new XmlDocument();
                XmlDoc.LoadXml(file.GetSourceString());
                FillNamespaces();
                XmlEl = (XmlElement)XmlDoc.SelectSingleNode("/w:numbering", Nsmgr);
            }
            catch (FileNotFoundException)
            {
                IsExist = false;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }

        }


        public AbstractNum GetAbstractNumByStyle(AbstractNumStyle style,bool createIfNotExist=false)
        {
            AbstractNum result = null;
            foreach (AbstractNum an in AbstartNums)
            {
                if(an.MultiLevelType!=style.MultiLevelType)
                    continue;
                Lvl firstLevel = an.FirstLvl;
                if (firstLevel.NumFmt.Value == style.Format && firstLevel.LvlText.Value == style.LvlText && firstLevel.LvlJc.Value == style.LvlHorizontalAlign &&
                    firstLevel.Ppr.Ind.Left == style.IndentingFirstLvl && firstLevel.Ppr.Ind.Hanging == style.Hanging)
                {
                    result = an;
                    break;
                }
            }

            if (result == null)
            {
                if (!createIfNotExist)
                    throw new AbstractNumNotFoundException("Не удалось найти AbstractNum по AbstractNumStyle");
                result = NewAbstractNum(style);
            }
            return result;
        }

        /// <summary>
        /// Перебирает связи Nums и возращает AbstractNumId
        /// </summary>
        /// <param name="numId"></param>
        /// <returns></returns>
        public AbstractNum GetAbstartNumByNumID(int numId)
        {
            bool NumIdFinded = false;
            foreach (Num num in Nums)
                if (num.NumId == numId)
                {
                    NumIdFinded = true;
                    foreach (AbstractNum an in AbstartNums)
                        if (an.AbstractNumId == num.AbstartNumId.Value)
                            return an;

                    throw new AbstractNumNotFoundException($"Не удалось найти w:abstractNum w:abstractNumId={num.AbstartNumId.Value}");
                }
            if (!NumIdFinded)
                throw new AbstractNumNotFoundException($"Не удалось найти w:num w:numId={numId} в файле word/numbering.xml");
            throw new AbstractNumNotFoundException($"Не удалось найти AbstractNum по numId={numId}");
        }

        public List<Num> Nums
        {
            get
            {
                return FindChilds<Num>();
            }
        }
        public List<AbstractNum> AbstartNums
        {
            get
            {
                return FindChilds<AbstractNum>();
            }
        }

        

        public AbstractNum NewAbstractNum(AbstractNumStyle style)
        {
            //Вычислить максимальный AbstractNum, заинкрементить и назначить новому классу
            int maxAbstractNumId = 0;
            foreach (AbstractNum an in AbstartNums)
                if (an.AbstractNumId > maxAbstractNumId)
                    maxAbstractNumId = an.AbstractNumId;

            AbstractNum result;
            if (AbstartNums.Count() > 0)
                result = NewNodeAfter<AbstractNum>(AbstartNums.Last().XmlEl);
            else
                result = NewNodeFirst<AbstractNum>();
            result.AbstractNumId = maxAbstractNumId + 1;

            result.SetAttribute("w15:restartNumberingAfterBreak", "0");
            result.NewNodeFirst<Nsid>().Value = GenerateGuid();
            result.NewNodeLast<MultiLevelType>().Value = style.MultiLevelType;
            result.NewNodeLast<Tmpl>().Value = GenerateGuid();
            for (int level=0;level<=(style.LvlCntMax??4);level++)
            {
                Lvl lvl = result.NewNodeLast<Lvl>();
                lvl.Ilvl = level;
                lvl.SetAttribute("w:tplc", GenerateGuid());
                lvl.Start.Value = 1;
                lvl.NumFmt.Value = style.Format;
                lvl.LvlText.Value = style.LvlText;
                lvl.LvlJc.Value = style.LvlHorizontalAlign;
                lvl.Ppr.Ind.Left = style.IndentingFirstLvl*(level+1);
                lvl.Ppr.Ind.Hanging = style.Hanging;
                lvl.Rpr.Font = style.FontName;
                if(style.FontSize!=null)
                lvl.Rpr.FontSize = style.FontSize;
            }
            return result;
        }

        public Num NewNum(int abstactNumId)
        {
            int maxNumId = 0;
            foreach (Num an in Nums)
                if (an.NumId > maxNumId)
                    maxNumId = an.NumId;

            Num result = NewNodeLast<Num>();
            result.NumId = maxNumId + 1;
            result.AbstartNumId.Value = abstactNumId;
            return result;
        }
    }

    public class AbstractNumStyle
    {
        public AbstractNumStyle(MULTI_LEVEL_TYPE multiLevelType, NUM_FMT format, string lvlText, HORIZONTAL_ALIGN lvlHorizontalAlign, Size indentingFirstLvl, Size hanging,string fontName,double? fontSize, int? lvlCntMax = null)
        {
            this.MultiLevelType = multiLevelType;
            this.Format = format;
            this.LvlText = lvlText;
            this.LvlHorizontalAlign = lvlHorizontalAlign;
            this.IndentingFirstLvl = indentingFirstLvl;
            this.Hanging = hanging;
            this.FontName = fontName; 
            this.FontSize = fontSize;
            this.LvlCntMax = lvlCntMax;
        }

        /// <summary>
        /// Тип списка 
        /// </summary>
        public MULTI_LEVEL_TYPE MultiLevelType;
        /// <summary>
        /// формат маркеров
        /// </summary>
        public NUM_FMT Format;
        /// <summary>
        /// маркер
        /// </summary>
        public string LvlText;
        /// <summary>
        /// выравнивание маркера
        /// </summary>
        public HORIZONTAL_ALIGN LvlHorizontalAlign;
        /// <summary>
        /// Отступ для первого уровня. Для остальных уровней высчитывается автоматически.
        /// </summary>
        public Size IndentingFirstLvl;

        public Size Hanging;
        /// <summary>
        /// Минимальное к-во уровней списка
        /// </summary>
        public int LvlCntMin;
        /// Максимальное к-во уровней списка
        public int? LvlCntMax;

        public string FontName;
        public double? FontSize;
    }

    public class AbstractNum : Node
    {
        public AbstractNum() : base("w:abstractNum") { }
        public AbstractNum(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:abstractNum") { }

        public AbstractNumStyle GetAbstractNumStyle()
        {
            return new AbstractNumStyle(MultiLevelType, FirstLvl.NumFmt.Value, FirstLvl.LvlText.Value,
                FirstLvl.LvlJc.Value, FirstLvl.Ppr.Ind.Left, FirstLvl.Ppr.Ind.Hanging,FirstLvl.Rpr.Font,FirstLvl.Rpr.FontSize, Levels.Count());
        }

        public MULTI_LEVEL_TYPE MultiLevelType
        {
            get { return MultiLevelNode.Value; }
            set { MultiLevelNode.Value = value; }
        }

        public MultiLevelType MultiLevelNode
        {
            get { return FindChildOrCreate<MultiLevelType>(); }
        }

        public int AbstractNumId
        {
            get
            {
                return Int32.Parse(XmlEl.GetAttribute("abstractNumId", XmlEl.NamespaceURI));
            }
            set
            {
                XmlEl.SetAttribute("abstractNumId",XmlEl.NamespaceURI, value.ToString());
            }
        }

        public Lvl FirstLvl
        {
            get
            {
                if (Levels.Count == 0)
                {
                    return NewNodeLast<Lvl>();
                }
                return Levels[0];
            }
        }

        public List<Lvl> Levels
        {
            get
            {
                return FindChilds<Lvl>();
            }
        }
        public Lvl GetLevel(int level)
        {
            return Levels.Where(x => x.Ilvl == level).FirstOrDefault();
        }
        

        public Lvl NewLevel(int level)
        {
            Lvl result = GetLevel(level);
            if (result != null)
                return result;
            result = NewNodeLast<Lvl>();
            result.Ilvl = level;
            return result;
        }
    }

    public class Num : Node
    {
        public Num() : base("w:num") { }
        public Num(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:num") { }
        public int NumId
        {
            get
            {
                return Int32.Parse(XmlEl.GetAttribute("numId", Nsmgr.LookupNamespace("w")));
            }
            set
            {
                XmlEl.SetAttribute("numId",XmlEl.NamespaceURI, value.ToString());
            }
        }
        public AbstractNumId AbstartNumId
        {
            get {
                var result = FindChild<AbstractNumId>();
                if(result==null)
                    result = NewNodeLast<AbstractNumId>();
                return result;
            }
        }
    }
    public class AbstractNumId : Node
    {
        public AbstractNumId() : base("w:abstractNumId") { }
        public AbstractNumId(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:abstractNumId") { }
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
    public class Lvl : Node
    {
        public Lvl() : base("w:lvl") { }
        public Lvl(Node parent) : base(parent, "w:lvl")
        {
            //<w:start w:val="1"/>
			//<w:numFmt w:val="bullet"/>
			//<w:lvlText w:val=""/>
			//<w:lvlJc w:val="left"/>
			//<w:pPr>
			//	<w:ind w:left="720" w:hanging="360"/>
			//</w:pPr>
			//<w:rPr>
			//	<w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/>
			//</w:rPr>
        }
        /// <summary>
        /// уровень
        /// </summary>
        public int Ilvl
        {
            get
            {
                return Int32.Parse(XmlEl.GetAttribute("w:ilvl"));
            }
            set
            {
                XmlEl.SetAttribute("ilvl", XmlEl.NamespaceURI, value.ToString());
            }
        }

        public Start Start
        {
            get
            {
                Start result = FindChild<Start>();
                if (result == null)
                    result = NewNodeFirst<Start>();
                return result;
            }
        }

        public NumFmt NumFmt
        {
            get
            {
                NumFmt result = FindChild<NumFmt>();
                if (result == null)
                    result = NewNodeLast<NumFmt>();
                return result;
            }
        }
        public LvlText LvlText
        {
            get
            {
                LvlText result = FindChild<LvlText>();
                if (result == null)
                    result = NewNodeLast<LvlText>();
                return result;
            }
        }
        public LvlJc LvlJc
        {
            get
            {
                LvlJc result = FindChild<LvlJc>();
                if (result == null)
                    result = NewNodeLast<LvlJc>();
                return result;
            }
        }
        public PProp Ppr
        {
            get
            {
                PProp result = FindChild<PProp>();
                if (result == null)
                    result = NewNodeLast<PProp>();
                return result;
            }
        }
        public RProp Rpr
        {
            get
            {
                RProp result = FindChild<RProp>();
                if (result == null)
                    result = NewNodeLast<RProp>();
                return result;
            }
        }
        public Lvl(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:lvl") { }
    }

    public class Start : Node
    {
        public Start() : base("w:start") { }
        public Start(Node parent) : base(parent, "w:start")
        { }
        public Start(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:start") { }
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


    public class LvlText : Node
    {
        public LvlText() : base("w:lvlText") { }
        public LvlText(Node parent) : base(parent, "w:lvlText")
        { }
        public LvlText(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:lvlText") { }
        public string Value
        {
            get
            {
                return XmlEl.GetAttribute("w:val");
            }
            set
            {
                XmlEl.SetAttribute("val", XmlEl.NamespaceURI, value);
            }
        }
    }
    public class Nsid : Node
    {
        public Nsid() : base("w:nsid") { }
        public Nsid(Node parent) : base(parent, "w:nsid")
        { }
        public Nsid(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:nsid") { }
        public string Value
        {
            get
            {
                return XmlEl.GetAttribute("w:val");
            }
            set
            {
                XmlEl.SetAttribute("val", XmlEl.NamespaceURI, value);
            }
        }
    }

    public enum MULTI_LEVEL_TYPE
    {
        /// <summary>
        /// определяет формат только на уровне
        /// </summary>
        SINGLE_LEVEL,
        /// <summary>
        /// список из нескольких уровней, каждый из которых имеет один и тот же вид (маркеры или уровень текста)
        /// </summary>
        MULTI_LEVEL,
        /// <summary>
        /// список из нескольких уровней, каждый из которых потенциально может быть разного типа (маркеры или уровень текста)
        /// </summary>
        HYBRID_MULTY_LEVEL
    }

    public class MultiLevelType : Node
    {
        public MultiLevelType() : base("w:multiLevelType") { }
        public MultiLevelType(Node parent) : base(parent, "w:multiLevelType")
        { }
        public MultiLevelType(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:multiLevelType") { }
        public MULTI_LEVEL_TYPE Value
        {
            get
            {
                return EnumExtentions.ToEnum<MULTI_LEVEL_TYPE>(GetAttribute("w:val"));
            }
            set
            {
                SetAttribute("w:val", value.ToStringValue());
            }
        }
    }

    public class Tmpl : Node
    {
        public Tmpl() : base("w:tmpl") { }
        public Tmpl(Node parent) : base(parent, "w:tmpl")
        { }
        public Tmpl(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:tmpl") { }
        public string Value
        {
            get
            {
                return XmlEl.GetAttribute("w:val");
            }
            set
            {
                XmlEl.SetAttribute("val", XmlEl.NamespaceURI, value);
            }
        }
    }

    public class LvlJc : Node
    {
        public LvlJc() : base("w:lvlJc") { }
        public LvlJc(Node parent) : base(parent, "w:lvlJc")
        { }
        public LvlJc(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:lvlJc") { }

        public HORIZONTAL_ALIGN Value
        {
            get
            {
                return EnumExtentions.ToEnum<HORIZONTAL_ALIGN>(GetAttribute("w:val"));
            }
            set
            {
                SetAttribute("w:val",value.ToStringValue());
            }
        }
    }
}
