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
            //桌面目录
            string desktopDir = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);

            //文件
            string Path_out = Path.Combine(desktopDir, "test_" + DateTime.Now.Ticks + ".docx");
            string tempFile = Path.Combine(desktopDir, "newtemplete.docx");      //获取模板路径，这个根据个人模板路径而定。

            WordDocument doc = new WordDocument(tempFile);
            
            //查找需要生成的节点的样式
            Aspose.Words.Lists.List numberList = null;
            ParagraphFormat paragraphFormat = null;
            NodeCollection nodes = doc.WordDoc.GetChildNodes(NodeType.Paragraph, true);
            foreach (Node node in nodes)
            {
                if (node.Range.Text.Contains("项目分解节点模板"))
                {
                    if (numberList == null)
                    {
                        numberList = ((Paragraph)node).ListFormat.List;
                        paragraphFormat = ((Paragraph)node).ParagraphFormat;
                    }

                    node.Remove();
                }
            }
            
            //替换封皮信息
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
            dic.Add("%<test>%", "测试替换...  ");
            doc.replaceAllWithBookmark(dic);

            //插入文档
            doc.insertDocumentAfterBookMark(new Document(Path.Combine(desktopDir, "ttt.docx")), "项目摘要", true);
            doc.insertDocumentAfterBookMark(new Document(Path.Combine(desktopDir, "ttt.docx")), "基本概念及内涵", true);
            doc.insertDocumentAfterBookMark(new Document(Path.Combine(desktopDir, "ttt.docx")), "军事需求分析", true);
            doc.insertDocumentAfterBookMark(new Document(Path.Combine(desktopDir, "ttt.docx")), "研究现状", true);
            doc.insertDocumentAfterBookMark(new Document(Path.Combine(desktopDir, "ttt.docx")), "研究目标", true);
            doc.insertDocumentAfterBookMark(new Document(Path.Combine(desktopDir, "ttt.docx")), "基础性问题", true);
            doc.insertDocumentAfterBookMark(new Document(Path.Combine(desktopDir, "ttt.docx")), "课题之间的关系", true);
            doc.insertDocumentAfterBookMark(new Document(Path.Combine(desktopDir, "ttt.docx")), "研究成果及考核指标", true);
            doc.insertDocumentAfterBookMark(new Document(Path.Combine(desktopDir, "ttt.docx")), "评估方案", true);
            doc.insertDocumentAfterBookMark(new Document(Path.Combine(desktopDir, "ttt.docx")), "预期效益", true);
            doc.insertDocumentAfterBookMark(new Document(Path.Combine(desktopDir, "ttt.docx")), "项目负责人C", true);
            doc.insertDocumentAfterBookMark(new Document(Path.Combine(desktopDir, "ttt.docx")), "研究团队", true);
            doc.insertDocumentAfterBookMark(new Document(Path.Combine(desktopDir, "ttt.docx")), "研究基础与保障条件", true);
            doc.insertDocumentAfterBookMark(new Document(Path.Combine(desktopDir, "ttt.docx")), "组织实施与风险控制", true);
            doc.insertDocumentAfterBookMark(new Document(Path.Combine(desktopDir, "ttt.docx")), "与有关计划关系", true);

            //输出项目分解的节点
            doc.WordDocBuilder.MoveToBookmark("项目分解详细");
            doc.WordDocBuilder.ListFormat.List = numberList;
            double oldFirstLineIndent = doc.WordDocBuilder.ParagraphFormat.FirstLineIndent;
            doc.WordDocBuilder.ParagraphFormat.FirstLineIndent = paragraphFormat.FirstLineIndent;
            doc.WordDocBuilder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
            doc.WordDocBuilder.Writeln("光刻机");
            doc.WordDocBuilder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;
            doc.WordDocBuilder.Writeln("、研究目标");
            doc.WordDocBuilder.StartBookmark("autoBookmark_1");
            doc.WordDocBuilder.EndBookmark("autoBookmark_1");

            doc.WordDocBuilder.Writeln("、研究内容");
            doc.WordDocBuilder.StartBookmark("autoBookmark_2");
            doc.WordDocBuilder.EndBookmark("autoBookmark_2");

            doc.WordDocBuilder.Writeln("、研究思路");
            doc.WordDocBuilder.StartBookmark("autoBookmark_3");
            doc.WordDocBuilder.EndBookmark("autoBookmark_3");

            doc.WordDocBuilder.Writeln("、负责单位及负责人");
            doc.WordDocBuilder.StartBookmark("autoBookmark_4");
            doc.WordDocBuilder.EndBookmark("autoBookmark_4");

            doc.WordDocBuilder.Writeln("、研究经费");
            doc.WordDocBuilder.StartBookmark("autoBookmark_5");
            doc.WordDocBuilder.EndBookmark("autoBookmark_5");

            doc.WordDocBuilder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
            doc.WordDocBuilder.Writeln("材料");
            doc.WordDocBuilder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;
            doc.WordDocBuilder.Writeln("、研究目标");
            doc.WordDocBuilder.StartBookmark("autoBookmark_6");
            doc.WordDocBuilder.EndBookmark("autoBookmark_6");

            doc.WordDocBuilder.Writeln("、研究内容");
            doc.WordDocBuilder.StartBookmark("autoBookmark_7");
            doc.WordDocBuilder.EndBookmark("autoBookmark_7");

            doc.WordDocBuilder.Writeln("、研究思路");
            doc.WordDocBuilder.StartBookmark("autoBookmark_8");
            doc.WordDocBuilder.EndBookmark("autoBookmark_8");

            doc.WordDocBuilder.Writeln("、负责单位及负责人");
            doc.WordDocBuilder.StartBookmark("autoBookmark_9");
            doc.WordDocBuilder.EndBookmark("autoBookmark_9");

            doc.WordDocBuilder.Writeln("、研究经费");
            doc.WordDocBuilder.StartBookmark("autoBookmark_10");
            doc.WordDocBuilder.EndBookmark("autoBookmark_10");

            doc.WordDocBuilder.ListFormat.RemoveNumbers();
            doc.WordDocBuilder.ParagraphFormat.FirstLineIndent = oldFirstLineIndent;
            //doc.WordDocBuilder.InsertHtml("<h1>1、分析数据</h1>");
            //doc.WordDocBuilder.InsertHtml("<h2>1.1 数据一</h2>");
            //doc.WordDocBuilder.InsertHtml("<h2>1.2 数据二</h2>");
            //doc.WordDocBuilder.InsertHtml("<h3>1.2.1 计算分析</h3>");
            //doc.WordDocBuilder.InsertHtml("<h1>2、分析数据</h1>");
            //doc.WordDocBuilder.InsertHtml("<h2>2.1 数据一</h2>");

            //填充课题详细内容
            for (int k = 1; k <= 10; k++)
            {
                doc.insertDocumentAfterBookMark(new Document(Path.Combine(desktopDir, "rrr.docx")), "autoBookmark_" + k, k == 10 ? true : false);
            }
            ////插入一个新页（横向）
            //doc.WordDocBuilder.MoveToBookmark("附件3");
            //doc.WordDocBuilder.InsertBreak(BreakType.SectionBreakNewPage);
            //doc.WordDocBuilder.PageSetup.Orientation = Aspose.Words.Orientation.Landscape;


            //统一编号
            doc.WordDoc.FirstSection.PageSetup.SectionStart = SectionStart.EvenPage;
            doc.WordDoc.FirstSection.PageSetup.RestartPageNumbering = true;

            //更新域
            doc.WordDoc.UpdateFields();
            doc.WordDoc.UpdateListLabels();
            doc.WordDoc.UpdatePageLayout();
            doc.WordDoc.UpdateTableLayout();
            doc.WordDoc.UpdateThumbnail();

            //保存word
            doc.WordDoc.Save(Path_out); //保存word

            //打开
            System.Diagnostics.Process.Start(Path_out);
        }
    }
}