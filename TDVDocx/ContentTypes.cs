using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace TDV.Docx
{
    public class ContentTypes:BaseNode
    {
        public ContentTypes(DocxDocument docx):base(docx)
        {
            docxDocument = docx;
            try
            {
                file = docx.sourceFolder.FindFile("[Content_Types].xml");

                xmlDoc = new XmlDocument();
                xmlDoc.LoadXml(file.GetSourceString());
                //nsmgr = new XmlNamespaceManager(xmlDoc.NameTable);
                //nsmgr.AddNamespace("Relationships", "http://schemas.openxmlformats.org/package/2006/relationships");
                FillNamespaces();
                xmlEl = (XmlElement)xmlDoc.SelectSingleNode(@"/DEFAULT:Types", nsmgr);
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
        //public Relationship NewRelationship(string target,RELATIONSIP_TYPE type)
        //{
        //    foreach (Relationship r in Relationships)
        //        if (r.Target == target)
        //            return r;
        //    Relationship newRel = NewNodeLast<Relationship>();
        //    newRel.Id = $"rId{GetMaxRelId() + 1}";
        //    newRel.Type = type;
        //    newRel.Target = target;
        //    return newRel;
        //}

       

        public List<Override> Overrides
        {
            get
            {
                return FindChilds<Override>();
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="partName">Путь к файлу абсолютный. пример /docProps/app.xml</param>
        /// <param name="createIfNotExist">после создания не забудьте заполнить ContentType</param>
        /// <returns></returns>
        public Override GetOverride(string partName, bool createIfNotExist = false)
        {
            Override result = null;
            foreach(Override o in Overrides)
                if (o.PartName == partName)
                { 
                    result = o;
                    break;
                }
            if(result==null)
            {
                if (!createIfNotExist)
                    throw new KeyNotFoundException($"Не найден Override c PartName={partName}");
                result = NewNodeLast<Override>();
                result.PartName = partName;
            }
            return result;
        }


    }

    public class Override : Node
    {
        public static class ContentTypes
        {
            public static string FOOTER = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml";
        }
        public Override() : base("Override") { }
        public Override(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "Override") { }
        public string PartName
        {
            get
            {
                return xmlEl.GetAttribute("PartName");
            }
            set
            {
                if (!value.StartsWith("/"))
                    value = "/" + value;
                xmlEl.SetAttribute("PartName", value);
            }
        }
        public string ContentType
        {
            get
            {
                return xmlEl.GetAttribute("ContentType");
            }
            set
            {
                xmlEl.SetAttribute("ContentType", value);
            }
        }
    }

}
