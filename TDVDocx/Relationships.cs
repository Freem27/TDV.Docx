using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace TDV.Docx
{
    public enum RELATIONSIP_TYPE
    {
        FOOTER, STYLES, ENDNOTES, NUMBERING, CUSTOM_XML, FOOTNOTES, WEB_SETTINGS, THEME, SETTINGS, FONT_TABLE, HEADER, HYPERLINK
    }

    public enum RELATIONSHIP_TARGET_MODE
    {
        NONE,EXTERNAL
    }
    public class WordRels:BaseNode
    {
        public WordRels(DocxDocument docx):base(docx)
        {
            DocxDocument = docx;
            try
            {
                file = docx.sourceFolder.FindFile("document.xml.rels", @"word/_rels");

                XmlDoc = new XmlDocument();
                XmlDoc.LoadXml(file.GetSourceString());
                //nsmgr = new XmlNamespaceManager(xmlDoc.NameTable);
                //nsmgr.AddNamespace("Relationships", "http://schemas.openxmlformats.org/package/2006/relationships");
                FillNamespaces();
                XmlEl = (XmlElement)XmlDoc.SelectSingleNode(@"/DEFAULT:Relationships", Nsmgr);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        /// <summary>
        /// вернет существующую, если такой target Уже есть
        /// иначе создаст новую
        /// </summary>
        /// <param name="target">путь к файлу отностилеьно document.xml.rels, например ../customXml/item1.xml</param>
        /// <returns></returns>
        public Relationship NewRelationship(string target,RELATIONSIP_TYPE type)
        {
            foreach (Relationship r in Relationships)
                if (r.Target == target)
                    return r;
            Relationship newRel = NewNodeLast<Relationship>();
            newRel.Id = $"rId{GetMaxRelId() + 1}";
            newRel.Type = type;
            newRel.Target = target;
            return newRel;
        }

        /// <summary>
        /// возвращает наибольший идентификатор связей Relationship
        /// </summary>
        /// <returns></returns>
        public int GetMaxRelId()
        {
            int result = 0;
            foreach(Relationship r in Relationships)
            {
                if (!r.Id.Contains("rId"))
                    continue;
                int curId = Int32.Parse(r.Id.Replace("rId", ""));
                if (curId > result)
                    result = curId;
            }
            return result;
        }

        public List<Relationship> Relationships
        {
            get
            {
                return FindChilds<Relationship>();
            }
        }

        public Relationship GetRelationshipById(string id)
        {
            foreach(Relationship r in Relationships)
                if(r.Id==id)
                {
                    return r;
                }
            throw new KeyNotFoundException($"Не найдена связь с id={id}");
        }

        internal ArchFile GetFileById(string id)
        {
            ArchFile result = null;
            string target = GetRelationshipById(id).Target;

            string filePath = Path.GetFullPath(Path.Combine(file.GetFolderPath(), target))
                .Substring(Directory.GetCurrentDirectory().Length + 1).Replace("\\", "/");
            string fileName = new FileInfo(filePath).Name;
            filePath = filePath.Replace(fileName, "");
            if (filePath.Last() == '/')
                filePath = filePath.Remove(filePath.Length - 1);
            //удалить имя файла
            result = DocxDocument.sourceFolder.FindFile(fileName, filePath);
            if (result == null)
                throw new FileNotFoundException($"Ну удалось найти файл с id={id}");
            return result;
        }

    }

    public class Relationship : Node
    {
        public Relationship() : base("Relationship") { }
        public Relationship(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "Relationship") { }
        public string Id
        {
            get
            {
                return XmlEl.GetAttribute("Id");
            }
            set
            {
                XmlEl.SetAttribute("Id", value);
            }
        }

        public RELATIONSIP_TYPE Type
        {
            get
            {
                switch (XmlEl.GetAttribute("Type"))
                {
                    case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer":
                        return RELATIONSIP_TYPE.FOOTER;
                    case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles":
                        return RELATIONSIP_TYPE.STYLES;
                    case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes":
                        return RELATIONSIP_TYPE.ENDNOTES;
                    case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering":
                        return RELATIONSIP_TYPE.NUMBERING;
                    case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml":
                        return RELATIONSIP_TYPE.CUSTOM_XML;
                    case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes":
                        return RELATIONSIP_TYPE.FOOTNOTES;
                    case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings":
                        return RELATIONSIP_TYPE.WEB_SETTINGS;
                    case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme":
                        return RELATIONSIP_TYPE.THEME;
                    case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings":
                        return RELATIONSIP_TYPE.SETTINGS;
                    case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable":
                        return RELATIONSIP_TYPE.FONT_TABLE;
                    case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header":
                        return RELATIONSIP_TYPE.HEADER;
                    case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink":
                        return RELATIONSIP_TYPE.HYPERLINK;
                }
                throw new Exception($"Неизвестный тип связи {XmlEl.GetAttribute("Type")}");
            }
            set
            {
                string type = "";
                switch (value)
                {
                    case RELATIONSIP_TYPE.FOOTER:
                        type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer";
                        break;
                    case RELATIONSIP_TYPE.STYLES:
                        type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
                        break;
                    case RELATIONSIP_TYPE.ENDNOTES:
                        type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes";
                        break;
                    case RELATIONSIP_TYPE.NUMBERING:
                        type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";
                        break;
                    case RELATIONSIP_TYPE.CUSTOM_XML:
                        type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml";
                        break;
                    case RELATIONSIP_TYPE.FOOTNOTES:
                        type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes";
                        break;
                    case RELATIONSIP_TYPE.WEB_SETTINGS:
                        type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings";
                        break;
                    case RELATIONSIP_TYPE.THEME:
                        type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
                        break;
                    case RELATIONSIP_TYPE.SETTINGS:
                        type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings";
                        break;
                    case RELATIONSIP_TYPE.FONT_TABLE:
                        type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable";
                        break;
                    case RELATIONSIP_TYPE.HEADER:
                        type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header";
                        break;
                    case RELATIONSIP_TYPE.HYPERLINK:
                        type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
                        TargetMode = RELATIONSHIP_TARGET_MODE.EXTERNAL;
                        break;
                    default:
                        throw new Exception($"Неизвестный тип связи {value.ToString()}");
                }                
                XmlEl.SetAttribute("Type", type);
            }
        }

        public string Target
        {
            get
            {
                return GetAttribute("Target");
            }
            set
            {
                if (string.IsNullOrEmpty(value))
                    RemoveAttribute("Target");
                else
                    SetAttribute("Target", value);
            }
        }

        public RELATIONSHIP_TARGET_MODE TargetMode
        {
            get
            {
                if (!HasAttribute("TargetMode"))
                    return RELATIONSHIP_TARGET_MODE.NONE;
                switch(GetAttribute("TargetMode"))
                {
                    case "External":
                        return RELATIONSHIP_TARGET_MODE.EXTERNAL;
                }
                throw new NotImplementedException();
            }
            set
            {
                switch (value)
                {
                    case RELATIONSHIP_TARGET_MODE.EXTERNAL:
                        SetAttribute("TargetMode", "External");
                        return;
                    case RELATIONSHIP_TARGET_MODE.NONE:
                        RemoveAttribute("TargetMode");
                        return;
                }
                throw new NotImplementedException();
            }
        }
    }

}
