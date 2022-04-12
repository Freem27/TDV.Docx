using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace TDV.Docx {
  public class ContentTypes : BaseNode {
    public ContentTypes(DocxDocument docx) : base(docx) {
      DocxDocument = docx;
      try {
        file = docx.sourceFolder.FindFile("[Content_Types].xml");

        XmlDoc = new XmlDocument();
        XmlDoc.LoadXml(file.GetSourceString());
        FillNamespaces();
        XmlEl = (XmlElement)XmlDoc.SelectSingleNode(@"/DEFAULT:Types", Nsmgr);
      }
      catch (Exception e) {
        Console.WriteLine(e.Message);
      }
    }

    public List<Override> Overrides {
      get {
        return FindChilds<Override>();
      }
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="partName">Путь к файлу абсолютный. пример /docProps/app.xml</param>
    /// <param name="createIfNotExist">после создания не забудьте заполнить ContentType</param>
    /// <returns></returns>
    public Override GetOverride(string partName, bool createIfNotExist = false) {
      Override result = null;
      foreach (Override o in Overrides)
        if (o.PartName == partName) {
          result = o;
          break;
        }
      if (result == null) {
        if (!createIfNotExist)
          throw new KeyNotFoundException($"Не найден Override c PartName={partName}");
        result = NewNodeLast<Override>();
        result.PartName = partName;
      }
      return result;
    }
  }

  public class Override : Node {
    public static class ContentTypes {
      public static string FOOTER = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml";
      public static string HEADER = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml";
      public static string COMMENTS = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml";
    }
    public Override() : base("Override") { }
    public Override(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "Override") { }
    public string PartName {
      get {
        return XmlEl.GetAttribute("PartName");
      }
      set {
        if (!value.StartsWith("/"))
          value = "/" + value;
        XmlEl.SetAttribute("PartName", value);
      }
    }
    public string ContentType {
      get {
        return XmlEl.GetAttribute("ContentType");
      }
      set {
        XmlEl.SetAttribute("ContentType", value);
      }
    }
  }
}