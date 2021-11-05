using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace TDV.Docx
{
    public class Styles: BaseNode
    {
        public Styles(DocxDocument docx):base(docx)
        {
            try
            {
                file = docx.sourceFolder.FindFile("styles.xml", @"word");

                xmlDoc = new XmlDocument();
                xmlDoc.LoadXml(file.GetSourceString());
                nsmgr = new XmlNamespaceManager(xmlDoc.NameTable);
                nsmgr.AddNamespace("w", xmlDoc.DocumentElement.NamespaceURI);
                xmlEl = (XmlElement)xmlDoc.SelectSingleNode(@"w:styles", nsmgr);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        public Style GetStyleById(string id)
        {
            return (Style)childNodes.Where(x => x is Style && ((Style)x).StyleId== id).FirstOrDefault();
        }
    }


    public class Style : Node
    {
        public Style() : base("w:style") { }
        public Style(Node parent) : base(parent, "w:style") { }
        public Style(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:style") {  }


        

        public T GetStyleProp<T>() where T : Node
        {
            T result = null;
            result = (T)childNodes.Where(x => x is T).FirstOrDefault();
            if (result == null)
            {
                if(basedOn!=null)
                    return basedOn.GetStyleProp<T>();
            }
            else
            {
                if (basedOn != null)
                {
                    T parentStyleProp= basedOn.GetStyleProp<T>();
                    if (parentStyleProp != null)
                        result.baseStyleNodes = parentStyleProp.childNodes;
                }
            }
            return result;
        }

        /// <summary>
        /// базовый стиль
        /// </summary>
        public Style basedOn
        {
            get
            {
                XmlElement el = (XmlElement)xmlEl.SelectSingleNode("w:basedOn", nsmgr);
                if (el == null)
                    return null;
                string baseSyleId = el.GetAttribute("w:val");
                return ((Styles) parent).GetStyleById(baseSyleId);
            }
        }

        public string StyleId
        {
            get { return xmlEl.GetAttribute("w:styleId"); }
        }

        public string StyleType
        {
            get { return xmlEl.GetAttribute("w:type"); }
        }

        /// <summary>
        /// <w:name w:val=\"Normal\"/>
        /// </summary>
        public string Name
        {
            get
            {
                XmlElement el = (XmlElement)xmlEl.SelectSingleNode("name", nsmgr);
                if (el == null)
                    return null;
                return el.GetAttribute("w:val");
            }
        }
    }
}
