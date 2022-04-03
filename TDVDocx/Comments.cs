using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace TDV.Docx {
    public class Comments : BaseNode {
        public Comments(DocxDocument docx) : base(docx) {
            DocxDocument = docx;
            try {
                file = docx.sourceFolder.FindFile("comments.xml");

                XmlDoc = new XmlDocument();
                XmlDoc.LoadXml(file.GetSourceString());
                FillNamespaces();
                XmlEl = (XmlElement)XmlDoc.SelectSingleNode(@"/w:comments", Nsmgr);
            }
            catch (FileNotFoundException) {
                ArchFolder wordFolder = docx.sourceFolder.GetFolder("word");
                file = wordFolder.AddFile($"comments.xml", new byte[0]);
                XmlDoc = new XmlDocument();
                XmlDoc.LoadXml(@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<w:comments xmlns:wpc=""http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"" xmlns:mc=""http://schemas.openxmlformats.org/markup-compatibility/2006"" xmlns:o=""urn:schemas-microsoft-com:office:office"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"" xmlns:m=""http://schemas.openxmlformats.org/officeDocument/2006/math"" xmlns:v=""urn:schemas-microsoft-com:vml"" xmlns:wp14=""http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"" xmlns:wp=""http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"" xmlns:w10=""urn:schemas-microsoft-com:office:word"" xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"" xmlns:w14=""http://schemas.microsoft.com/office/word/2010/wordml"" xmlns:w15=""http://schemas.microsoft.com/office/word/2012/wordml"" xmlns:wpg=""http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"" xmlns:wpi=""http://schemas.microsoft.com/office/word/2010/wordprocessingInk"" xmlns:wne=""http://schemas.microsoft.com/office/word/2006/wordml"" xmlns:wps=""http://schemas.microsoft.com/office/word/2010/wordprocessingShape"" mc:Ignorable=""w14 w15 wp14"">
	
</w:comments>");
                FillNamespaces();
                XmlEl = (XmlElement)XmlDoc.SelectSingleNode(@"/w:comments", Nsmgr);
                Override ov = docx.ContentTypes.GetOverride(file.GetFullPath(), true);
                ov.ContentType = Override.ContentTypes.COMMENTS;

                if (!docx.WordRels.Relationships.Where(x => x.Target == "comments.xml").Any())
                    docx.WordRels.NewRelationship("comments.xml", RELATIONSIP_TYPE.COMMENT);
            }
            catch (Exception e) {
                Console.WriteLine(e.Message);
            }
        }

        public List<Comment> CommentsList {
            get {
                return FindChilds<Comment>();
            }
        }

        public Comment GetCommentById(int id, bool createIfNotExist = false) {
            Comment result = CommentsList.Where(x => x.Id == id).FirstOrDefault();
            if (result == null) {
                if (!createIfNotExist)
                    throw new KeyNotFoundException();
                result = NewComment(id);
            }
            return result;
        }

        public Comment NewComment(int id, string author = "TDV", string initials = "") {
            Comment result = NewNodeLast<Comment>();
            result.Id = id;
            result.Author = author;
            result.Date = DateTime.Now;
            result.Initials = initials;
            return result;
        }

    }

    public class Comment : Node {
        public Comment() : base("w:comment") { }
        public Comment(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:comment") { }
        public override string Text {
            get {
                return base.Text;
            }
            set {
                while (Paragraphs.Count > 0)
                    Paragraphs.First().Delete();
                Paragraph p = NewNodeLast<Paragraph>();
                p.Text = value;
            }
        }

        public List<Paragraph> Paragraphs {
            get {
                return FindChilds<Paragraph>();
            }
        }
        public int Id {
            get {
                return Int32.Parse(GetAttribute("w:id"));
            }
            set {
                SetAttribute("w:id", value.ToString());
            }

        }

        public DateTime? Date {
            get {
                try {
                    return DateTime.Parse(GetAttribute("w:date"));
                }
                catch {
                    return null;
                }
            }
            set {
                if (value == null)
                    RemoveAttribute("w:date");
                else
                    SetAttribute("w:date", ((DateTime)value).ToString("yyyy-MM-ddTHH:mm:ssZ"));
            }
        }

        /// <summary>
        /// Автор комментария
        /// </summary>
        public string Author {
            get {
                return GetAttribute("w:author");
            }
            set {
                SetAttribute("w:author", value);
            }
        }

        /// <summary>
        /// Инициалы автора
        /// </summary>
        public string Initials {
            get {
                return GetAttribute("w:initials");
            }
            set {
                SetAttribute("w:initials", value);
            }
        }
    }

    public class CommentRangeStart : Node {
        public CommentRangeStart() : base("w:commentRangeStart") { }
        public CommentRangeStart(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:commentRangeStart") { }

        public string Comment {
            get {
                DocxDocument docx = GetDocxDocument();
                return docx.Comments.GetCommentById(Id)?.Text;
            }
            set {
                DocxDocument docx = GetDocxDocument();
                docx.Comments.GetCommentById(Id, true).Text = value;
            }
        }

        public string CommentedText {
            get {
                string result = "";
                Node next = NextNodeRecurcieve;
                while (next != null) {
                    if (next is R) {
                        result += next.Text;
                    }
                    else if (next is Paragraph) {
                        result += "\n";
                    }
                    else if (next is CommentRangeEnd && ((CommentRangeEnd)next).Id == Id)
                        break;
                    next = next.NextNodeRecurcieve;
                }
                return result;
            }
        }

        public int Id {
            get {
                return Int32.Parse(GetAttribute("w:id"));
            }
            set {
                SetAttribute("w:id", value.ToString());
            }
        }

        public override void InitXmlElement() {
            base.InitXmlElement();
            Id = GetDocxDocument().Document.GetNextId();
        }
    }

    public class CommentRangeEnd : Node {
        public CommentRangeEnd() : base("w:commentRangeEnd") { }
        public CommentRangeEnd(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:commentRangeEnd") { }
        public int Id {
            get {
                return Int32.Parse(GetAttribute("w:id"));
            }
            set {
                SetAttribute("w:id", value.ToString());
                //добавить следующей нодой Run с commentReference
                if (!(NextNode is R)) {
                    Parent.NewNodeAfter<R>(this);
                }
                ((R)NextNode).FindChildOrCreate<CommentReference>().Id = Id;
            }
        }
    }

    public class CommentReference : Node {
        public CommentReference() : base("w:commentReference") { }
        public CommentReference(XmlElement xmlElement, Node parent) : base(xmlElement, parent, "w:commentReference") { }
        public int Id {
            get {
                return Int32.Parse(GetAttribute("w:id"));
            }
            set {
                SetAttribute("w:id", value.ToString());
            }
        }
    }

    public static class RNodesExtensions {
        /// <summary>
        /// Уставнавливает комментарий для последовательно идущих нод
        /// ноды должны быть из одного параграфа и располагаться друг за другом
        /// </summary>
        /// <param name="rList"></param>
        public static void SetComment(this List<R> rList, string commmentText, string author = "TDV") {
            if (rList == null || rList.Count() < 1) {
                throw new ArgumentNullException("rList == null || rList.Count() < 1");
            }

            Node parent = rList.First().Parent;
            DocxDocument docxDocument = rList.First().GetDocxDocument();

            if (parent is Ins || parent is Del) {
                parent = parent.Parent;
            }

            CommentRangeStart commentRangeStart = parent.NewNodeBefore<CommentRangeStart>(rList.First());
            CommentRangeEnd commentRangeEnd = parent.NewNodeAfter<CommentRangeEnd>(rList.Last());
            commentRangeEnd.Id = commentRangeStart.Id;
            docxDocument.Comments.NewComment(commentRangeStart.Id, author).Text = commmentText;
        }
    }
}
