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
        }

        //设置合同模板全局，增加上下左右边距等
        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            //设置上下左右边距
            WordApp.ActiveDocument.PageSetup.TopMargin = 60;
            WordApp.ActiveDocument.PageSetup.BottomMargin = 60;
            WordApp.ActiveDocument.PageSetup.LeftMargin = 80;
            WordApp.ActiveDocument.PageSetup.RightMargin = 80;

            //全局大纲变为“正文”
            WordApp.ActiveDocument.Paragraphs.Format.OutlineLevel = WdOutlineLevel.wdOutlineLevelBodyText;
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            Style style = null;
            //添加和设置正文格式
            try
            {
                style = WordApp.ActiveDocument.Styles["正文"];
            }
            catch(Exception)
            {
                WordApp.ActiveDocument.Styles.Add("正文");
                style = WordApp.ActiveDocument.Styles["正文"];
            }
            style.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevelBodyText;     //大纲“正文”文本
            style.ParagraphFormat.CharacterUnitFirstLineIndent = 2;                         //首行缩进两个字符
            style.ParagraphFormat.LineSpacing = 24;                                         //行距24磅

            style.Font.NameFarEast = "仿宋";
            style.Font.NameAscii = "Times New Roman";
            style.Font.NameOther = "Times New Roman";
            style.Font.Name = "正文";

        }
       
    }
}
