using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace TDV.Docx
{
    public class Theme:BaseNode
    {
        internal Theme(DocxDocument docx,ArchFile file):base(docx)
        {
            DocxDocument = docx;
            try
            {
                this.file = file;
                XmlDoc = new XmlDocument();
                XmlDoc.LoadXml(file.GetSourceString());
                FillNamespaces();
                XmlEl = (XmlElement)XmlDoc.SelectSingleNode(@"/a:theme", Nsmgr);
            }
            catch(FileNotFoundException)
            {
                IsExist = false;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        public string Name
        {
            get
            {
                return GetAttribute("name");
            }
            set
            {
                SetAttribute("name",value);
            }
        }

        public string GetMajorFont()
        {
            return this.ThemeElements?.FontScheme?.MajorFont?.Latin?.TypeFace;
        }

        public ThemeElements ThemeElements
        {
            get { return FindChild<ThemeElements>(); }
        }

        public bool IsObjectDefaults
        {
            get
            {
                return FindChild<ObjectDefaults>() != null;
            }
            set
            {
                if (value)
                    FindChildOrCreate<ObjectDefaults>();
                else
                    FindChild<ObjectDefaults>()?.Delete();
            }
        }

    }

    public class ThemeElements : Node
    {
        public ThemeElements() : base("a:themeElements") { }
        public ThemeElements(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "a:themeElements") { }

        public FontScheme FontScheme
        {
            get
            {
                return FindChild<FontScheme>();
            }
        }

        public ClrScheme ClrScheme
        {
            get { return FindChildOrCreate<ClrScheme>(); }
        }
    }

    public class ClrScheme : Node
    {
        public ClrScheme() : base("a:clrScheme") { }
        public ClrScheme(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "a:clrScheme") { }


    }
    public class SrgbClr : Node
    {
        public SrgbClr() : base("a:srgbClr") { }
        public SrgbClr(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "a:srgbClr") { }

        /// STRING
        public string Value
        {
            get
            {
                if (HasAttribute("val"))
                    return GetAttribute("val");
                else return null;
            }
            set
            {
                if (string.IsNullOrEmpty(value))
                    RemoveAttribute("val");
                else
                    SetAttribute("val", value);
            }
        }
    }

    public class ObjectDefaults : Node
    {
        public ObjectDefaults() : base("a:objectDefaults") { }
        public ObjectDefaults(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "a:objectDefaults") { }
    }

    public class FontScheme : Node
    {
        public FontScheme() : base("a:fontScheme") { }
        public FontScheme(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "a:fontScheme") { }

        public MajorFont MajorFont
        {
            get { return FindChild<MajorFont>(); }
        }
    }

    public class MajorFont : Node
    {
        public MajorFont() : base("a:majorFont") { }
        public MajorFont(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "a:majorFont") { }

        public Latin Latin
        {
            get { return FindChild<Latin>(); }
        }
    }
    public class Latin : Node
    {
        public Latin() : base("a:latin") { }
        public Latin(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "a:latin") { }

        public string TypeFace
        {
            get
            {
                if (HasAttribute("typeface"))
                    return GetAttribute("typeface");
                else return null;
            }
            set
            {
                SetAttribute("typeface", value);
            }
        }
    }
}
