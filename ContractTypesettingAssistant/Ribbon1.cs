using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Word;
using System.Windows.Forms;

namespace ContractTypesettingAssistant
{
    public partial class Ribbon1
    {
        private Microsoft.Office.Interop.Word.Application WordApp;
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            WordApp = Globals.ThisAddIn.Application;
            create_style_zhengwen();
        }

        //设置合同模板全局，增加上下左右边距等
        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            //设置上下左右边距
            WordApp.ActiveDocument.PageSetup.TopMargin = 70;
            WordApp.ActiveDocument.PageSetup.BottomMargin = 60;
            WordApp.ActiveDocument.PageSetup.LeftMargin = 80;
            WordApp.ActiveDocument.PageSetup.RightMargin = 80;

            //全局大纲变为“正文”
            WordApp.ActiveDocument.Paragraphs.Format.OutlineLevel = WdOutlineLevel.wdOutlineLevelBodyText;
            WordApp.ActiveDocument.SelectAllEditableRanges();
            WordApp.Selection.Find.Text = ":";
            WordApp.Selection.Find.Replacement.Text = "：+";
        }

        //设置主标题格式
        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            //WordApp.Selection.set_Style("合同主标题");
        }

        //创建正文文本格式
        private void create_style_zhengwen()
        {
            Style style = null;
            //添加和设置正文格式
            try
            {
                style = WordApp.ActiveDocument.Styles["正文"];
            }
            catch (Exception)
            {
                WordApp.ActiveDocument.Styles.Add("正文");
                style = WordApp.ActiveDocument.Styles["正文"];
            }
            style.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevelBodyText;     //大纲“正文”文本
            style.ParagraphFormat.FirstLineIndent = 5;
            style.ParagraphFormat.CharacterUnitFirstLineIndent = 2;                         //首行缩进两个字符
            style.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;       //行距设置为固定值
            style.ParagraphFormat.LineSpacing = 24;                                         //行距24磅

            style.Font.NameFarEast = "仿宋";                                                //中文字体
            style.Font.NameAscii = "仿宋";                                       //英文格式
            style.Font.NameOther = "仿宋";                                       //字符格式
            style.Font.Name = "仿宋";                                                       //格式名称
            style.Font.Size = 12;

        }

        //设置正文格式
        private void button5_Click(object sender, RibbonControlEventArgs e)
        {
            WordApp.Selection.ClearFormatting();
            WordApp.Selection.set_Style("正文");
        }
    }
}
