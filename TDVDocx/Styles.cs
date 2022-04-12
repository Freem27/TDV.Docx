using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace TDV.Docx {
  public enum STYLE_TYPE { TABLE, CHARACTER, PARAGRAPH, NUMBERING }
  public class Styles : BaseNode {
    public Styles(DocxDocument docx) : base(docx) {
      try {
        file = docx.sourceFolder.FindFile("styles.xml", @"word");

        XmlDoc = new XmlDocument();
        XmlDoc.LoadXml(file.GetSourceString());
        FillNamespaces();
        XmlEl = (XmlElement)XmlDoc.SelectSingleNode(@"w:styles", Nsmgr);
      }
      catch (Exception e) {
        Console.WriteLine(e.Message);
      }
    }

    public List<Style> StylesList {
      get { return FindChilds<Style>(); }
    }

    public DocDefaults DocDefaults {
      get {
        return FindChild<DocDefaults>();
      }
    }

    public Style GetDefaultParagraphFontStyle() {
      Style result = null;
      result = StylesList.Where(x => x.Name == "Default Paragraph Font").FirstOrDefault();
      if (result == null) {
        result = NewNodeLast<Style>();
        result.StyleId = $"a{GetMaxStyleId() + 1}";
        result.Name = "Default Paragraph Font";
        result.Default = "1";
        result.IsSemiHidden = true;
        result.IsUnhideWhenUsed = true;
      }
      return result;
    }

    public Style GetStyleByName(string name) {
      return StylesList.Where(x => x.Name.ToLower() == name.ToLower()).FirstOrDefault();
    }
    public int GetMaxStyleId(string idPrefix = null) {
      int maxId = 0;
      foreach (Style s in StylesList) {
        if ((idPrefix != null && s.StyleId.StartsWith(idPrefix)) || idPrefix == null) {
          if (Int32.TryParse(idPrefix != null ? s.StyleId.Replace(idPrefix, "") : s.StyleId, out int styleIdInt))
            if (styleIdInt > maxId)
              maxId = styleIdInt;
        }
      }
      return maxId;
    }

    public Style GetStyleById(string id) {
      return StylesList.Where(x => x.StyleId == id).FirstOrDefault();
    }
  }

  public class Style : Node {
    public Style() : base("w:style") { }
    public Style(Node parent) : base(parent, "w:style") { }
    public Style(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:style") { }

    public bool IsSemiHidden {
      get {
        return FindChild<SemiHidden>() != null;
      }
      set {
        if (value)
          FindChildOrCreate<SemiHidden>();
        else
          FindChild<SemiHidden>()?.Delete();
      }
    }

    public bool IsUnhideWhenUsed {
      get {
        return FindChild<UnhideWhenUsed>() != null;
      }
      set {
        if (value)
          FindChildOrCreate<UnhideWhenUsed>();
        else
          FindChild<UnhideWhenUsed>()?.Delete();
      }
    }

    public T GetStyleProp<T>() where T : Node {
      T result = null;
      result = (T)ChildNodes.Where(x => x is T).FirstOrDefault();
      if (result == null) {
        if (basedOn != null)
          return basedOn.GetStyleProp<T>();
      }
      else {
        if (basedOn != null) {
          T parentStyleProp = basedOn.GetStyleProp<T>();
          if (parentStyleProp != null)
            result.baseStyleNodes = parentStyleProp.ChildNodes;
        }
      }
      return result;
    }

    /// <summary>
    /// базовый стиль
    /// </summary>
    public Style basedOn {
      get {
        BasedOn basedOn = FindChild<BasedOn>();
        if (basedOn == null)
          return null;
        return ((Styles)Parent).GetStyleById(basedOn.Value);
      }
      set {
        FindChildOrCreate<BasedOn>().Value = value.StyleId;
      }
    }

    public string StyleId {
      get {
        try {
          return GetAttribute("w:styleId");
        }
        catch (KeyNotFoundException) {
          return null;
        }
      }
      set {
        if (string.IsNullOrEmpty(value))
          RemoveAttribute("w:styleId");
        else
          SetAttribute("w:styleId", value);
      }
    }
    public string Default {
      get {
        try {
          return GetAttribute("w:default");
        }
        catch (KeyNotFoundException) {
          return null;
        }
      }
      set {
        if (string.IsNullOrEmpty(value))
          RemoveAttribute("w:default");
        else
          SetAttribute("w:default", value);
      }
    }

    public STYLE_TYPE Type {
      get {
        switch (GetAttribute("w:type")) {
          case "character":
            return STYLE_TYPE.CHARACTER;
          case "paragraph":
            return STYLE_TYPE.PARAGRAPH;
          case "table":
            return STYLE_TYPE.TABLE;
          case "numbering":
            return STYLE_TYPE.NUMBERING;
        }
        throw new NotImplementedException();
      }
      set {
        switch (value) {
          case STYLE_TYPE.CHARACTER:
            SetAttribute("w:type", "character");
            return;
          case STYLE_TYPE.PARAGRAPH:
            SetAttribute("w:type", "paragraph");
            return;
          case STYLE_TYPE.TABLE:
            SetAttribute("w:type", "table");
            return;
          case STYLE_TYPE.NUMBERING:
            SetAttribute("w:type", "numbering");
            return;
        }
        throw new NotImplementedException();
      }
    }

    public string Name {
      get {
        return FindChild<Name>()?.Value ?? null;
      }
      set {
        FindChildOrCreate<Name>().Value = value;
      }
    }


    public int UiPriority {
      get {
        return FindChild<UiPriority>().Value;
      }
      set {
        FindChildOrCreate<UiPriority>().Value = value;
      }
    }
  }

  public class DocDefaults : Node {
    public DocDefaults() : base("w:docDefaults") { }
    public DocDefaults(Node parent) : base(parent, "w:docDefaults") { }
    public DocDefaults(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:docDefaults") { }

    public RPrDefault RPrDefault {
      get { return FindChild<RPrDefault>(); }
    }

    public PPrDefault PPrDefault {
      get { return FindChild<PPrDefault>(); }
    }
  }

  public class PPrDefault : Node {
    public PPrDefault() : base("w:pPrDefault") { }
    public PPrDefault(Node parent) : base(parent, "w:pPrDefault") { }
    public PPrDefault(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:pPrDefault") { }

    public PProp PProp {
      get { return FindChild<PProp>(); }
    }
  }

  public class RPrDefault : Node {
    public RPrDefault() : base("w:rPrDefault") { }
    public RPrDefault(Node parent) : base(parent, "w:rPrDefault") { }
    public RPrDefault(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:rPrDefault") { }

    public RProp RProp {
      get { return FindChild<RProp>(); }
    }

  }

  public class SemiHidden : Node {
    public SemiHidden() : base("w:semiHidden") { }
    public SemiHidden(Node parent) : base(parent, "w:semiHidden") { }
    public SemiHidden(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:semiHidden") { }

  }

  public class UnhideWhenUsed : Node {
    public UnhideWhenUsed() : base("w:unhideWhenUsed") { }
    public UnhideWhenUsed(Node parent) : base(parent, "w:unhideWhenUsed") { }
    public UnhideWhenUsed(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:unhideWhenUsed") { }
  }

  public class Name : Node {
    public Name() : base("w:name") { }
    public Name(Node parent) : base(parent, "w:name") { }
    public Name(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:name") { }

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
  public class BasedOn : Node {
    public BasedOn() : base("w:basedOn") { }
    public BasedOn(Node parent) : base(parent, "w:basedOn") { }
    public BasedOn(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:basedOn") { }

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

  public class UiPriority : Node {
    public UiPriority() : base("w:uiPriority") { }
    public UiPriority(Node parent) : base(parent, "w:uiPriority") { }
    public UiPriority(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:uiPriority") { }

    public int Value {
      get {
        return Int32.Parse(GetAttribute("w:val"));
      }
      set {
        SetAttribute("w:val", value.ToString());
      }
    }
  }

  public class RStyleNode : Node {
    public RStyleNode() : base("w:rStyle") { }
    public RStyleNode(Node parent) : base(parent, "w:rStyle") { }
    public RStyleNode(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:rStyle") { }

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
}