using Aspose.Words;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace WordTest
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            string desktopDir = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);

            string Path_out = Path.Combine(desktopDir, "test_" + DateTime.Now.Ticks + ".docx");
            string tempFile = Path.Combine(desktopDir, "newtemplete.docx");      //获取模板路径，这个根据个人模板路径而定。

            WordDocument doc = new WordDocument(tempFile);

            Dictionary<string, object> dic = new Dictionary<string, object>();   //创建键值对   第一个string 为书签名称 第二个string为要填充的数据
            dic.Add("首页密级", "保密");
            dic.Add("申报领域", "芯片技术");
            dic.Add("申报方向", "量子芯片技术");
            dic.Add("项目名称", "量子芯片制造");
            dic.Add("单位名称", "芯片技术研究院");
            dic.Add("单位常用名", "芯片技术研究院天津分院");
            dic.Add("项目负责人", "张三");
            dic.Add("单位联系人", "张五");
            dic.Add("联系电话", "68111111");
            dic.Add("通信地址", "天津滨海区");
            dic.Add("研究周期", "5");
            dic.Add("研究经费", "1000");
            dic.Add("项目关键字", "aaa;bbb;ccc;ddd");
            
            doc.insertAllWithBookmark(dic);

            dic = new Dictionary<string, object>();
            dic.Add("%<test>%", "测试替换...");
            doc.replaceAllWithBookmark(dic);

            doc.insertDocumentAfterBookMark(new Document(Path.Combine(desktopDir, "ttt.docx")), "基本概念及内涵");
            doc.insertDocumentAfterBookMark(new Document(Path.Combine(desktopDir, "ttt.docx")), "军事需求分析");
            doc.insertDocumentAfterBookMark(new Document(Path.Combine(desktopDir, "ttt.docx")), "研究现状");
            doc.insertDocumentAfterBookMark(new Document(Path.Combine(desktopDir, "ttt.docx")), "研究目标");
            doc.insertDocumentAfterBookMark(new Document(Path.Combine(desktopDir, "ttt.docx")), "基础性问题");

            doc.WordDoc.UpdateFields();// 更新域
            doc.WordDoc.UpdateListLabels();
            doc.WordDoc.UpdatePageLayout();
            doc.WordDoc.UpdateTableLayout();
            doc.WordDoc.UpdateThumbnail();

            // Set the appended document to appear on the next page.
            doc.WordDoc.FirstSection.PageSetup.SectionStart = SectionStart.EvenPage;
            // Restart the page numbering for the document to be appended.
            doc.WordDoc.FirstSection.PageSetup.RestartPageNumbering = true;

            // Go to the primary footer
            //doc.DocBuilder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
            // Add fields for current page number
            //doc.DocBuilder.InsertField("PAGE", "");
            // Add any custom text
            //doc.DocBuilder.Write(" / ");
            //// Add field for total page numbers in document
            //doc.DocBuilder.InsertField("NUMPAGES", "");

            doc.WordDocBuilder.MoveToBookmark("附件3");
            doc.WordDocBuilder.InsertBreak(BreakType.SectionBreakNewPage);
            doc.WordDocBuilder.PageSetup.Orientation = Aspose.Words.Orientation.Landscape;
                        
            //doc.DocBuilder.InsertHtml("<h1>1、分析数据</h1>");
            //doc.DocBuilder.InsertHtml("<h2>1.1 数据一</h2>");
            //doc.DocBuilder.InsertHtml("<h2>1.2 数据二</h2>");
            //doc.DocBuilder.InsertHtml("<h3>1.2.1 计算分析</h3>");

            //doc.DocBuilder.InsertHtml("<h1>2、分析数据</h1>");
            //doc.DocBuilder.InsertHtml("<h2>2.1 数据一</h2>");

            //Aspose.Words.Lists.List numberedList = doc.WordDoc.Lists.Add(Aspose.Words.Lists.ListTemplate.NumberDefault);
            //numberedList.ListLevels[1].NumberStyle = NumberStyle.Arabic;
            //numberedList.ListLevels[1].NumberFormat = "\x0000.\x0001";

            doc.WordDocBuilder.MoveToBookmark("项目分解详细");
            Aspose.Words.Lists.List numberList = null;
            NodeCollection nodes = doc.WordDoc.GetChildNodes(NodeType.Paragraph, true);
            foreach (Node node in nodes)
            {
               
            }

            doc.WordDoc.Save(Path_out); //保存word

            //打开
            System.Diagnostics.Process.Start(Path_out);
        }
    }
}