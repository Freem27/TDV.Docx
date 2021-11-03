using System;
using System.IO;
using System.IO.Compression;
using System.Collections.Generic;
using System.Text;

namespace TDV.Docx
{
    public class DocxDocument
    {
        internal ArchFolder sourceFolder;
        public Document document;
        public WordRels wordRels;
        public Styles styles;
        public FootNotes footNotes;
        public Numbering numbering;
        public DocxDocument(Stream stream)
        {
            sourceFolder = new ArchFolder(null);
            ZipArchive arch = new ZipArchive(stream);
            foreach (ZipArchiveEntry e in arch.Entries)
            {
                using (Stream s = e.Open())
                {
                    using (MemoryStream ms = new MemoryStream())
                    {
                        s.CopyTo(ms);
                        byte[] content = ms.ToArray();
                        sourceFolder.AddFile(e.FullName, content);
                    }
                }
            }
            document = new Document(this);
            wordRels = new WordRels(this);
            styles = new Styles(this);
            footNotes = new FootNotes(this);
            numbering = new Numbering(this);
        }
        public static string Test()
        {
            return "test";
        }

        public MemoryStream ToStream()
        {
            MemoryStream result = new MemoryStream();
            using (ZipArchive arch = new ZipArchive(result, ZipArchiveMode.Create, true))
            {
                var t = sourceFolder.GetAllFilesRecurcive();
                foreach (var item in t)
                {
                    var newEntry = arch.CreateEntry(item.Key);
                    using (Stream sw = newEntry.Open())
                    {
                        sw.Write(item.Value.content, 0, item.Value.content.Length);
                    }
                }
            }
            result.Seek(0, SeekOrigin.Begin);
            return result;
        }

        public void Apply()
        {
            document.Apply();
            if(footNotes.IsExist)
                footNotes.Apply();
            if(numbering.IsExist)
                numbering.Apply();
            
        }

        public byte[] ToBytes()
        {
            return ToStream().ToArray();
        }
    }

    internal class ArchFolder
    {
        public string Name;
        public ArchFolder parent;

        public ArchFile FindFile(string fileName, string path = null)
        {
            string searchPath = Path.Combine(path ?? "", fileName).Replace("\\",@"/");
            ArchFile result = null;
            foreach (var item in GetAllFilesRecurcive())
            {
                if (!string.IsNullOrEmpty(path) && item.Value.GetFullPath() == searchPath)
                {
                    result = item.Value;
                    break;
                }
                else if (string.IsNullOrEmpty(path) && item.Value.Name == fileName)
                {
                    result = item.Value;
                    break;
                }
            }

            if (result == null)
                throw new FileNotFoundException($"Не удалось найти файл {searchPath} в контейнере докумена");
            return result;
        }
        public ArchFolder(string name, ArchFolder parent = null)
        {
            Name = name;
            this.parent = parent;
            folders = new Dictionary<string, ArchFolder>();
            files = new Dictionary<string, ArchFile>();
        }

        public void AddFile(string filePath, byte[] content)
        {
            string[] pathList = filePath.Split('/');
            ArchFolder targetFolder = this;
            for (int i = 0; i < pathList.Length; i++)
            {
                string entryName = pathList[i];
                if (i == pathList.Length - 1)
                {
                    targetFolder.files.Add(entryName, new ArchFile(entryName, content, targetFolder));
                }
                else
                {
                    if (!targetFolder.folders.ContainsKey(entryName))
                        targetFolder.folders.Add(entryName, new ArchFolder(entryName, targetFolder));
                    targetFolder = targetFolder.folders[entryName];
                }
            }
        }

        public List<KeyValuePair<string, ArchFile>> GetAllFilesRecurcive()
        {
            List<KeyValuePair<string, ArchFile>> result = new List<KeyValuePair<string, ArchFile>>();
            foreach (ArchFile file in files.Values)
                result.Add(new KeyValuePair<string, ArchFile>(file.GetFullPath(), file));
            foreach (ArchFolder folder in folders.Values)
            {
                result.AddRange(folder.GetAllFilesRecurcive()); // folder.GetAllFilesRecurcive()
            }
            return result;
        }
        public override string ToString()
        {
            return Name;
        }
        Dictionary<string, ArchFolder> folders;
        Dictionary<string, ArchFile> files;
    }

    internal class ArchFile
    {
        public byte[] content;
        public string Name;
        public ArchFolder parent;
        public ArchFile(string name, byte[] content, ArchFolder parent = null)
        {
            this.content = content;
            this.Name = name;
            this.parent = parent;
        }

        public override string ToString()
        {
            return Name;
        }

        public string GetSourceString()
        {
            return Encoding.UTF8.GetString(content);
        }

        public string GetFolderPath()
        {
            if (parent == null)
                return "";
            return parent.parent.Name;
        }

        public string GetFullPath()
        {
            string result = Name;
            ArchFolder p = parent;
            while (p != null && p.Name != null)
            {
                result = $"{p.Name}/{result}";
                p = p.parent;
            }
            return result;
        }
    }
}
