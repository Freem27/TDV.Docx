using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace TDV.Docx {
    public enum RELATIONSHIP_TARGET_MODE {
        NONE, EXTERNAL
    }
    public class WordRels : BaseNode {
        public WordRels(DocxDocument docx) : base(docx) {
            DocxDocument = docx;
            try {
                file = docx.sourceFolder.FindFile("document.xml.rels", @"word/_rels");

                XmlDoc = new XmlDocument();
                XmlDoc.LoadXml(file.GetSourceString());
                //nsmgr = new XmlNamespaceManager(xmlDoc.NameTable);
                //nsmgr.AddNamespace("Relationships", "http://schemas.openxmlformats.org/package/2006/relationships");
                FillNamespaces();
                XmlEl = (XmlElement)XmlDoc.SelectSingleNode(@"/DEFAULT:Relationships", Nsmgr);
            }
            catch (Exception e) {
                Console.WriteLine(e.Message);
            }
        }

        /// <summary>
        /// вернет существующую, если такой target Уже есть
        /// иначе создаст новую
        /// </summary>
        /// <param name="target">путь к файлу отностилеьно document.xml.rels, например ../customXml/item1.xml</param>
        /// <returns></returns>
        public Relationship NewRelationship(string target, RELATIONSIP_TYPE type) {
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
        public int GetMaxRelId() {
            int result = 0;
            foreach (Relationship r in Relationships) {
                if (!r.Id.Contains("rId"))
                    continue;
                int curId = Int32.Parse(r.Id.Replace("rId", ""));
                if (curId > result)
                    result = curId;
            }
            return result;
        }

        public List<Relationship> Relationships {
            get {
                return FindChilds<Relationship>();
            }
        }

        public Relationship GetRelationshipById(string id) {
            foreach (Relationship r in Relationships)
                if (r.Id == id) {
                    return r;
                }
            throw new KeyNotFoundException($"Не найдена связь с id={id}");
        }

        internal ArchFile GetFileById(string id) {
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
                throw new FileNotFoundException($"Не удалось найти файл с id={id}");
            return result;
        }
    }

    public class Relationship : Node {
        public Relationship() : base("Relationship") { }
        public Relationship(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "Relationship") { }
        public string Id {
            get {
                return XmlEl.GetAttribute("Id");
            }
            set {
                XmlEl.SetAttribute("Id", value);
            }
        }

        public RELATIONSIP_TYPE Type {
            get {
                return EnumExtentions.ToEnum<RELATIONSIP_TYPE>(GetAttribute("Type"));
            }
            set {
                XmlEl.SetAttribute("Type", value.ToStringValue());
            }
        }

        public string Target {
            get {
                return GetAttribute("Target");
            }
            set {
                if (string.IsNullOrEmpty(value))
                    RemoveAttribute("Target");
                else
                    SetAttribute("Target", value);
            }
        }

        public RELATIONSHIP_TARGET_MODE TargetMode {
            get {
                if (!HasAttribute("TargetMode"))
                    return RELATIONSHIP_TARGET_MODE.NONE;
                switch (GetAttribute("TargetMode")) {
                    case "External":
                        return RELATIONSHIP_TARGET_MODE.EXTERNAL;
                }
                throw new NotImplementedException();
            }
            set {
                switch (value) {
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