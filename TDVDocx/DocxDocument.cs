using System;
using System.IO;
using System.IO.Compression;
using System.Collections.Generic;
using System.Text;

namespace TDV.Docx {
  public class DocxDocument {
    internal ArchFolder sourceFolder;
    public Document Document;
    public WordRels WordRels;
    public Styles Styles;
    public ContentTypes ContentTypes;
    public Settings Settings;
    private Comments _comments;
    public List<Theme> ThemesList;

    public Theme ThemeDefault {
      get {
        foreach (Theme t in ThemesList) {
          if (t.IsObjectDefaults) {
            return t;
          }
        }
        return null;
      }
    }
    public Comments Comments {
      get {
        if (_comments == null)
          _comments = new Comments(this);
        return _comments;
      }
    }
    private FootNotes _footNotes;
    private Numbering _numbering;
    public FootNotes FootNotes {
      get {
        if (_footNotes == null)
          _footNotes = new FootNotes(this);
        return _footNotes;
      }
    }
    public Numbering Numbering {
      get {
        if (_numbering == null)
          _numbering = new Numbering(this);
        return _numbering;
      }
    }
    public List<BaseNode> FilesForApply;
    internal Dictionary<string, Header> headers;
    internal Dictionary<string, Footer> footers;

    public Header GetHeader(string id) {
      if (!headers.ContainsKey(id)) {
        ArchFile file = WordRels.GetFileById(id);
        headers.Add(id, new Header(this, file, WordRels.GetRelationshipById(id)));
      }
      return headers[id];
    }
    public Footer GetFooter(string id) {
      if (!footers.ContainsKey(id)) {
        ArchFile file = WordRels.GetFileById(id);
        footers.Add(id, new Footer(this, file, WordRels.GetRelationshipById(id)));
      }
      return footers[id];
    }




    public DocxDocument(Stream stream) {
      headers = new Dictionary<string, Header>();
      footers = new Dictionary<string, Footer>();
      FilesForApply = new List<BaseNode>();
      sourceFolder = new ArchFolder(null);
      ZipArchive arch = new ZipArchive(stream);

      foreach (ZipArchiveEntry e in arch.Entries) {
        using (Stream s = e.Open()) {
          using (MemoryStream ms = new MemoryStream()) {
            s.CopyTo(ms);
            byte[] content = ms.ToArray();
            sourceFolder.AddFile(e.FullName, content);
          }
        }
      }

      Document = new Document(this);
      WordRels = new WordRels(this);
      Styles = new Styles(this);
      ContentTypes = new ContentTypes(this);
      Settings = new Settings(this);
      ThemesList = new List<Theme>();
      foreach (ArchFile file in sourceFolder.GetFolder("word").GetFolder("theme").GetFiles()) {
        ThemesList.Add(new Theme(this, file));
      }
    }

    public MemoryStream ToStream() {
      MemoryStream result = new MemoryStream();
      using (ZipArchive arch = new ZipArchive(result, ZipArchiveMode.Create, true)) {
        var t = sourceFolder.GetAllFilesRecurcive();
        foreach (var item in t) {
          var newEntry = arch.CreateEntry(item.Key);
          using (Stream sw = newEntry.Open()) {
            sw.Write(item.Value.Content, 0, item.Value.Content.Length);
          }
        }
      }
      result.Seek(0, SeekOrigin.Begin);
      return result;
    }

    public void Apply() {
      foreach (BaseNode f in FilesForApply)
        if (f.IsExist)
          f.Apply();
    }

    public void ApplyAllFixes() {
      foreach (BaseNode f in FilesForApply)
        f.ApplyAllFixes();
    }

    public byte[] ToBytes() {
      return ToStream().ToArray();
    }
  }

  internal class ArchFolder {
    public string Name;
    public ArchFolder Parent;

    public List<ArchFile> GetFiles() {
      List<ArchFile> result = new List<ArchFile>();
      foreach (string fileName in files.Keys) {
        result.Add(files[fileName]);
      }
      return result;
    }

    public ArchFolder GetFolder(string name, bool createIfExist = false) {
      ArchFolder result = null;

      if (folders.ContainsKey(name))
        result = folders[name];
      else {
        if (!createIfExist)
          throw new FileNotFoundException($"Не удалось найти папку {name}");
        folders.Add(name, new ArchFolder(name, this));
        result = folders[name];
      }
      return result;
    }

    public ArchFile FindFile(string fileName, string path = null) {
      string searchPath = Path.Combine(path ?? "", fileName).Replace("\\", @"/");
      ArchFile result = null;
      foreach (var item in GetAllFilesRecurcive()) {
        if (!string.IsNullOrEmpty(path) && item.Value.GetFullPath() == searchPath) {
          result = item.Value;
          break;
        }
        else if (string.IsNullOrEmpty(path) && item.Value.Name == fileName) {
          result = item.Value;
          break;
        }
      }

      if (result == null)
        throw new FileNotFoundException($"Не удалось найти файл {searchPath} в контейнере докумена");
      return result;
    }

    public ArchFolder(string name, ArchFolder parent = null) {
      Name = name;
      this.Parent = parent;
      folders = new Dictionary<string, ArchFolder>();
      files = new Dictionary<string, ArchFile>();
    }

    public ArchFile AddFile(string filePath, byte[] content) {
      ArchFile result = null;
      string[] pathList = filePath.Split('/');
      ArchFolder targetFolder = this;
      for (int i = 0; i < pathList.Length; i++) {
        string entryName = pathList[i];
        if (i == pathList.Length - 1) {
          result = new ArchFile(entryName, content, targetFolder);
          targetFolder.files.Add(entryName, result);
        }
        else {
          if (!targetFolder.folders.ContainsKey(entryName)) {
            targetFolder.folders.Add(entryName, new ArchFolder(entryName, targetFolder));
          }
          targetFolder = targetFolder.folders[entryName];
        }
      }
      return result;
    }

    public List<KeyValuePair<string, ArchFile>> GetAllFilesRecurcive() {
      List<KeyValuePair<string, ArchFile>> result = new List<KeyValuePair<string, ArchFile>>();
      foreach (ArchFile file in files.Values)
        result.Add(new KeyValuePair<string, ArchFile>(file.GetFullPath(), file));
      foreach (ArchFolder folder in folders.Values) {
        result.AddRange(folder.GetAllFilesRecurcive()); // folder.GetAllFilesRecurcive()
      }
      return result;
    }
    public override string ToString() {
      return Name;
    }
    Dictionary<string, ArchFolder> folders;
    Dictionary<string, ArchFile> files;
  }

  internal class ArchFile {
    public byte[] Content;
    public string Name;
    public ArchFolder Parent;
    public ArchFile(string name, byte[] content, ArchFolder parent = null) {
      this.Content = content;
      this.Name = name;
      this.Parent = parent;
    }

    public override string ToString() {
      return Name;
    }

    public string GetSourceString() {
      return Encoding.UTF8.GetString(Content);
    }

    public string GetFolderPath() {
      if (Parent == null)
        return "";
      return Parent.Parent.Name;
    }

    public string GetFullPath() {
      string result = Name;
      ArchFolder p = Parent;
      while (p != null && p.Name != null) {
        result = $"{p.Name}/{result}";
        p = p.Parent;
      }
      return result;
    }
  }
}