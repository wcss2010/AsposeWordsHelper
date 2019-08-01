using Aspose.Words.Drawing;
using Aspose.Words.Replacing;
using Aspose.Words.Tables;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
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
        public DocumentBuilder WordDocBuilder
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
            foreach (KeyValuePair<string, object> kvp in dict)   //循环键值对
            {
                //移动书签到指定位置
                WordDocBuilder.MoveToBookmark(kvp.Key);  //将光标移入书签的位置

                //填充数值
                if (kvp.Value != null)
                {
                    if (kvp.Value.GetType().FullName == typeof(System.Drawing.Image).FullName || kvp.Value.GetType().FullName == typeof(System.Drawing.Bitmap).FullName)
                    {
                        documentBuilder.InsertImage((System.Drawing.Image)kvp.Value);
                    }
                    else
                    {
                        WordDocBuilder.Write(kvp.Value.ToString());
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
        /// <param name="isNeedDeleteEnter"></param>
        /// <returns></returns>
        public Document insertDocumentAfterBookMark(Document tobeInserted, string bookmark,bool isNeedDeleteEnter)
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
                    Node insertAfterNode = documentBuilder.CurrentParagraph.PreviousSibling;
                    insertDocumentAfterNode(insertAfterNode, tobeInserted);
                }
            }
            else
            {
                // if bookmark is not provided, add the document at the end
                appendDoc(tobeInserted);
            }

            if (isNeedDeleteEnter)
            {
                if (WordDocBuilder.CurrentParagraph != null)
                {
                    WordDocBuilder.CurrentParagraph.Remove();
                }
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

        /// <summary>
        /// 合并单元格
        /// </summary>
        /// <param name="startCell">开始单元格</param>
        /// <param name="endCell">结束单元格</param>
        /// <param name="table">表格</param>
        public void mergeCells(Cell startCell, Cell endCell, Table table)
        {
            Point startCellPos = new Point(startCell.ParentRow.IndexOf(startCell), table.IndexOf(startCell.ParentRow));
            Point endCellPos = new Point(endCell.ParentRow.IndexOf(endCell), table.IndexOf(endCell.ParentRow));
            Rectangle mergeRange = new Rectangle(System.Math.Min(startCellPos.X,
                    endCellPos.X), System.Math.Min(startCellPos.Y, endCellPos.Y),
                    System.Math.Abs(endCellPos.X - startCellPos.X) + 1,
                    System.Math.Abs(endCellPos.Y - startCellPos.Y) + 1);
            foreach (Row row in table.Rows)
            {
                foreach (Cell cell in row.Cells)
                {
                    Point currentPos = new Point(row.IndexOf(cell), table.IndexOf(row));
                    if (mergeRange.Contains(currentPos))
                    {
                        if (currentPos.X == mergeRange.X)
                        {
                            cell.CellFormat.HorizontalMerge = (CellMerge.First);
                        }
                        else
                        {
                            cell.CellFormat.HorizontalMerge = (CellMerge.Previous);
                        }
                        if (currentPos.Y == mergeRange.Y)
                        {
                            cell.CellFormat.VerticalMerge = (CellMerge.First);
                        }
                        else
                        {
                            cell.CellFormat.VerticalMerge = (CellMerge.Previous);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 添加DataTable数据到Table对象（必须是已经有表头的表格）
        /// </summary>
        /// <param name="table"></param>
        /// <param name="table"></param>
        public void fillDataToTable(Table table, DataTable dataTable)
        {
            //获得列数
            int cellCount = table.Rows[table.Rows.Count - 1].Cells.Count;
            //填充数据
            foreach (DataRow dr in dataTable.Rows)
            {
                //创建新行
                Aspose.Words.Tables.Row rowObj = new Row(table.Document);
                for (int k = 0; k < dataTable.Columns.Count; k++)
                {
                    if (k >= cellCount)
                    {
                        continue;
                    }

                    //创建列
                    Aspose.Words.Tables.Cell cellObj = new Cell(table.Document);
                    Aspose.Words.Paragraph p = new Paragraph(table.Document);
                    p.AppendChild(new Run(table.Document, dr[k] != null ? dr[k].ToString() : string.Empty));
                    cellObj.AppendChild(p);
                    rowObj.Cells.Add(cellObj);
                }

                //保存列
                if (rowObj.Cells.Count >= 1)
                {
                    table.Rows.Add(rowObj);
                }
            }
        }

        /// <summary>
        /// 引入表格从一个另外的Word文档
        /// </summary>
        /// <param name="insertAfterNode"></param>
        /// <param name="table"></param>
        public void importTableAfterNode(Node insertAfterNode, Table table)
        {
            Document tableDoc = new Document();
            tableDoc.FirstSection.Body.AppendChild(new NodeImporter(table.Document, tableDoc, ImportFormatMode.KeepSourceFormatting).ImportNode(table, true));
            insertDocumentAfterNode(insertAfterNode, tableDoc);
        }
        
        /// <summary>
        /// 插入水印文本
        /// </summary>
        /// <param name="fontName">字体名称</param>
        /// <param name="fontSize">字号</param>
        /// <param name="width">文字矩阵宽度</param>
        /// <param name="height">文字矩阵高度</param>
        /// <param name="rotation">文字矩阵偏转角度</param>
        /// <param name="fillColor">文字矩阵填充颜色</param>
        /// <param name="watermarkText">水印文本</param>
        public void insertWatermarkText(string fontName, double fontSize, int width, int height, int rotation, Color fillColor, string watermarkText)
        {
            // Create a watermark shape. This will be a WordArt shape. 
            // You are free to try other shape types as watermarks.
            Shape watermark = new Shape(WordDoc, ShapeType.TextPlainText);

            // Set up the text of the watermark.
            watermark.TextPath.Text = watermarkText;
            watermark.TextPath.FontFamily = fontName;
            watermark.TextPath.Size = fontSize;
            watermark.Width = width;
            watermark.Height = height;
            // Text will be directed from the bottom-left to the top-right corner.
            watermark.Rotation = rotation;
            // Remove the following two lines if you need a solid black text.
            watermark.Fill.Color = fillColor; // Try LightGray to get more Word-style watermark
            watermark.StrokeColor = fillColor; // Try LightGray to get more Word-style watermark

            // Place the watermark in the page center.
            watermark.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            watermark.RelativeVerticalPosition = RelativeVerticalPosition.Page;
            watermark.WrapType = WrapType.None;
            watermark.VerticalAlignment = VerticalAlignment.Center;
            watermark.HorizontalAlignment = HorizontalAlignment.Center;

            // Create a new paragraph and append the watermark to this paragraph.
            Paragraph watermarkPara = new Paragraph(WordDoc);
            watermarkPara.AppendChild(watermark);

            // Insert the watermark into all headers of each document section.
            foreach (Section sect in WordDoc.Sections)
            {
                // There could be up to three different headers in each section, since we want
                // the watermark to appear on all pages, insert into all headers.
                insertWatermarkIntoHeader(watermarkPara, sect, HeaderFooterType.HeaderPrimary);
                insertWatermarkIntoHeader(watermarkPara, sect, HeaderFooterType.HeaderFirst);
                insertWatermarkIntoHeader(watermarkPara, sect, HeaderFooterType.HeaderEven);
            }
        }

        /// <summary>
        /// 在页头插入水印文本
        /// </summary>
        /// <param name="watermarkPara"></param>
        /// <param name="sect"></param>
        /// <param name="headerType"></param>
        public void insertWatermarkIntoHeader(Paragraph watermarkPara, Section sect, HeaderFooterType headerType)
        {
            HeaderFooter header = sect.HeadersFooters[headerType];

            if (header == null)
            {
                // There is no header of the specified type in the current section, create it.
                header = new HeaderFooter(sect.Document, headerType);
                sect.HeadersFooters.Add(header);
            }

            // Insert a clone of the watermark into the header.
            header.AppendChild(watermarkPara.Clone(true));
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