using Aspose.Words.Replacing;
using System;
using System.Collections.Generic;
using System.Text;

namespace Aspose.Words
{
    /// <summary>
    /// 基于Aspose.Words的扩展
    /// </summary>
    public class WordDocument
    {
        private Document document = null;
        /// <summary>
        /// Word文档
        /// </summary>
        public Document WordDoc
        {
            get { return document; }
        }

        private DocumentBuilder documentBuilder = null;
        /// <summary>
        /// Word文档操作
        /// </summary>
        public DocumentBuilder DocBuilder
        {
            get { return documentBuilder; }
        }

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="docFile">Word文档</param>
        public WordDocument(string docFile)
        {
            document = new Document(docFile);
            documentBuilder = new DocumentBuilder(document);
        }

        /// <summary>
        /// 插入字典中所有的内容(String或Image)
        /// </summary>
        /// <param name="dict">要插入内容的字典（Key=书签,Value=内容）</param>
        public void insertAllWithBookmark(Dictionary<string, object> dict)
        {
            foreach (KeyValuePair<string,object> kvp in dict)   //循环键值对
            {
                //移动书签到指定位置
                DocBuilder.MoveToBookmark(kvp.Key);  //将光标移入书签的位置

                //填充数值
                if (kvp.Value != null)
                {
                    if (kvp.Value.GetType().FullName == typeof(System.Drawing.Image).FullName || kvp.Value.GetType().FullName == typeof(System.Drawing.Bitmap).FullName)
                    {
                        documentBuilder.InsertImage((System.Drawing.Image)kvp.Value);
                    }
                    else
                    {
                        DocBuilder.Write(kvp.Value.ToString());
                    }
                }
            }
        }

        /// <summary>
        /// 替换字典中所有的内容(String或Image)
        /// </summary>
        /// <param name="dict">要替换内容的字典（Key=书签,Value=内容）</param>
        public void replaceAllWithBookmark(Dictionary<string, object> dict)
        {
            foreach (KeyValuePair<string, object> kvp in dict)   //循环键值对
            {
                //填充数值
                if (kvp.Value != null)
                {
                    if (kvp.Value.GetType().FullName == typeof(System.Drawing.Image).FullName || kvp.Value.GetType().FullName == typeof(System.Drawing.Bitmap).FullName)
                    {
                        System.Text.RegularExpressions.Regex reg = new System.Text.RegularExpressions.Regex(kvp.Key);
                        WordDoc.Range.Replace(reg, new ReplaceAndInsertImage((System.Drawing.Image)kvp.Value), false);
                    }
                    else
                    {
                        WordDoc.Range.Replace(kvp.Key, kvp.Value.ToString(), false, false);
                    }
                }
            }
        }

        /// <summary>
        /// 插入文档到书签后
        /// </summary>
        /// <param name="tobeInserted"></param>
        /// <param name="bookmark"></param>
        /// <returns></returns>
        public Document insertDocumentAfterBookMark(Document tobeInserted, string bookmark)
        {
            // check to be inserted doc
            if (tobeInserted == null)
            {
                return WordDoc;
            }
                        
            // check bookmark and then process
            if (bookmark != null && bookmark.Trim().Length > 0)
            {
                Bookmark bm = WordDoc.Range.Bookmarks[bookmark];
                if (bm != null)
                {
                    documentBuilder.MoveToBookmark(bookmark);
                    documentBuilder.Writeln();
                    Node insertAfterNode = documentBuilder.CurrentParagraph.PreviousSibling;
                    insertDocumentAfterNode(insertAfterNode, tobeInserted);
                }
            }
            else
            {
                // if bookmark is not provided, add the document at the end
                appendDoc(tobeInserted);
            }
            return WordDoc;
        }

        /// <summary>
        /// 插入文档到节点后
        /// </summary>
        /// <param name="insertAfterNode"></param>
        /// <param name="srcDoc"></param>
        public void insertDocumentAfterNode(Node insertAfterNode, Document srcDoc)
        {
            // Make sure that the node is either a pargraph or table.
            if ((insertAfterNode.NodeType != NodeType.Paragraph)
            & (insertAfterNode.NodeType != NodeType.Table))
                throw new Exception("The destination node should be either a paragraph or table.");

            //We will be inserting into the parent of the destination paragraph.
            CompositeNode dstStory = insertAfterNode.ParentNode;
            //Remove empty paragraphs from the end of document
            while (null != srcDoc.LastSection.Body.LastParagraph && !srcDoc.LastSection.Body.LastParagraph.HasChildNodes)
            {
                srcDoc.LastSection.Body.LastParagraph.Remove();
            }

            NodeImporter importer = new NodeImporter(srcDoc, WordDoc, ImportFormatMode.KeepSourceFormatting);
            //Loop through all sections in the source document.
            int sectCount = srcDoc.Sections.Count;
            for (int sectIndex = 0; sectIndex < sectCount; sectIndex++)
            {
                Section srcSection = srcDoc.Sections[sectIndex];
                //Loop through all block level nodes (paragraphs and tables) in the body of the section.
                int nodeCount = srcSection.Body.ChildNodes.Count;
                for (int nodeIndex = 0; nodeIndex < nodeCount; nodeIndex++)
                {
                    Node srcNode = srcSection.Body.ChildNodes[nodeIndex];
                    Node newNode = importer.ImportNode(srcNode, true);
                    dstStory.InsertAfter(newNode, insertAfterNode);
                    insertAfterNode = newNode;
                }
            }
        }
        
        /// <summary>
        /// 添加到文档
        /// </summary>
        /// <param name="srcDoc"></param>
        /// <param name="includeSection"></param>
        public void appendDoc(Document srcDoc, bool includeSection)
        {
            // Loop through all sections in the source document.
            // Section nodes are immediate children of the Document node so we can
            // just enumerate the Document.
            if (includeSection)
            {
                foreach (Section srcSection in srcDoc.Sections)
                {
                    Node dstNode = WordDoc.ImportNode(srcSection, true, ImportFormatMode.UseDestinationStyles);
                    WordDoc.AppendChild(dstNode);
                }
            }
            else
            {
                //find the last paragraph of the last section
                Node node = WordDoc.LastSection.Body.LastParagraph;

                if (node == null)
                {
                    node = new Paragraph(srcDoc);
                    WordDoc.LastSection.Body.AppendChild(node);
                }

                if ((node.NodeType != NodeType.Paragraph)
                & (node.NodeType != NodeType.Table))
                {
                    throw new Exception("Use appendDoc(dstDoc, srcDoc, true) instead of appendDoc(dstDoc, srcDoc, false)");
                }
                insertDocumentAfterNode(node, srcDoc);
            }
        }

        /// <summary>
        /// 添加到文档
        /// </summary>
        /// <param name="srcDoc"></param>
        public void appendDoc(Document srcDoc)
        {
            appendDoc(srcDoc, true);
        }
    }

    public class ReplaceAndInsertImage : IReplacingCallback
    {
        /// <summary>
        /// 需要插入的图片
        /// </summary>
        public System.Drawing.Image ImageObj { get; set; }

        public ReplaceAndInsertImage(System.Drawing.Image img)
        {
            this.ImageObj = img;
        }

        public ReplaceAction Replacing(ReplacingArgs e)
        {
            //获取当前节点
            var node = e.MatchNode;
            //获取当前文档
            Document doc = node.Document as Document;
            DocumentBuilder builder = new DocumentBuilder(doc);
            //将光标移动到指定节点
            builder.MoveTo(node);
            //插入图片
            builder.InsertImage(ImageObj);
            return ReplaceAction.Replace;
        }
    }
}