



using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace TDV.Docx
{
    public class WordRels
    {
        private ArchFile file;
        XmlNamespaceManager nsmgr;
        private XmlDocument xmlDoc;
        XmlElement xmlEl;
        public DocxDocument docxDocument;
        public WordRels(DocxDocument docx)
        {
            docxDocument = docx;
            try
            {
                file = docx.sourceFolder.FindFile("document.xml.rels", @"word/_rels");

                xmlDoc = new XmlDocument();
                xmlDoc.LoadXml(file.GetSourceString());
                nsmgr = new XmlNamespaceManager(xmlDoc.NameTable);
                nsmgr.AddNamespace("Relationships", "http://schemas.openxmlformats.org/package/2006/relationships");
                xmlEl = (XmlElement)xmlDoc.SelectSingleNode(@"/Relationships:Relationships", nsmgr);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        internal ArchFile GetFileById(string id)
        {
            ArchFile result = null;
            string target = null;
            foreach (XmlElement rel in xmlEl.ChildNodes)
            {
                if (rel.GetAttribute("Id") == id)
                {
                    target = rel.GetAttribute("Target");
                    break;
                }
            }

            if (target == null)
                throw new KeyNotFoundException($"Не найдена связь с id={id} в файле {file.GetFullPath()}");

            string filePath = Path.GetFullPath(Path.Combine(file.GetFolderPath(), target))
                .Substring(Directory.GetCurrentDirectory().Length + 1).Replace("\\", "/");
            string fileName = new FileInfo(filePath).Name;
            filePath = filePath.Replace(fileName, "");
            if (filePath.Last() == '/')
                filePath = filePath.Remove(filePath.Length - 1);
            //удалить имя файла
            result = docxDocument.sourceFolder.FindFile(fileName, filePath);
            /*if (!target.Contains(@"/"))
            {
                result = docxDocument.sourceFolder.FindFile(target, file.GetFolderPath());
            }
            else
            {
                
                throw new Exception("Не реализована обработка путей файлов в значении атрибута Target");
            }*/

            if (result == null)
                throw new FileNotFoundException($"Ну удалось найти файл с id={id}");
            return result;
        }

    }


    public class Header : BaseNode
    {
        private ArchFile file;

        private XmlDocument xmlDoc;
        internal Header(ArchFile file) : base("w:hdr")
        {
            this.file = file;
            xmlDoc = new XmlDocument();
            
            xmlDoc.LoadXml(file.GetSourceString());
            nsmgr = new XmlNamespaceManager(xmlDoc.NameTable);
            nsmgr.AddNamespace("w", xmlDoc.DocumentElement.NamespaceURI);
            xmlEl = (XmlElement)xmlDoc.SelectSingleNode("/w:hdr", nsmgr);
        }

        public string Text()
        {
            string result = string.Join(" ", childNodes.Where(x => x is Paragraph).Select(x => ((Paragraph)x).Text));
            return result;
        }

        public void Apply()
        {
            using (StringWriter stringWriter = new StringWriter())
            using (XmlWriter xw = XmlWriter.Create(stringWriter))
            {
                xmlDoc.WriteTo(xw);
                xw.Flush();
                file.content = Encoding.UTF8.GetBytes(stringWriter.GetStringBuilder().ToString());
            }
        }
    }
}
