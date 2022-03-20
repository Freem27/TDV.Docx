using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace TDV.Docx {
    public class Settings : BaseNode {
        public Settings(DocxDocument docx) : base(docx) {
            DocxDocument = docx;
            try {
                file = docx.sourceFolder.FindFile("settings.xml", "word");

                XmlDoc = new XmlDocument();
                XmlDoc.LoadXml(file.GetSourceString());
                FillNamespaces();
                XmlEl = (XmlElement)XmlDoc.SelectSingleNode(@"/w:settings", Nsmgr);
            }
            catch (Exception e) {
                Console.WriteLine(e.Message);
            }
        }

        public Rsids Rsids {
            get {
                return FindChildOrCreate<Rsids>();
            }
        }

        public Rsid AppenndRsid() {
            return Rsids.NewNodeLast<Rsid>();
        }
    }

    public class Rsids : Node {
        public Rsids() : base("w:rsids") { }
        public Rsids(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:rsids") { }

        public RsidRoot RsidRoot {
            get {
                return FindChildOrCreate<RsidRoot>();
            }
        }
        public List<Rsid> RsidsList {
            get {
                return FindChilds<Rsid>();
            }
        }
    }

    public class RsidRoot : Node {
        public RsidRoot() : base("w:rsidRoot") { }
        public RsidRoot(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:rsidRoot") { }

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

        public override void InitXmlElement() {
            base.InitXmlElement();
            Value = GenerateGuid();
        }
    }

    public class Rsid : Node {
        public Rsid() : base("w:rsid") { }
        public Rsid(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:rsid") { }

        public override void InitXmlElement() {
            base.InitXmlElement();
            Value = GenerateGuid();
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
    }
}