﻿using System;
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
            //创建正文
            create_style_zhengwen();
            create_style_biaoti();

            //设置上下左右边距
            WordApp.ActiveDocument.PageSetup.TopMargin = 70;
            WordApp.ActiveDocument.PageSetup.BottomMargin = 60;
            WordApp.ActiveDocument.PageSetup.LeftMargin = 80;
            WordApp.ActiveDocument.PageSetup.RightMargin = 80;

            //全局大纲变为“正文”
            WordApp.ActiveDocument.Paragraphs.Format.OutlineLevel = WdOutlineLevel.wdOutlineLevelBodyText;

            //把英文冒号替换为中文冒号
            WordApp.Selection.Find.Execute(":",
                null,
                null,
                null,
                null,
                null,
                null,
                null,
                null,
                "：",
                WdReplace.wdReplaceAll,
                null,
                null,
                null,
                null);
        }


        //设置正文格式
        private void button5_Click(object sender, RibbonControlEventArgs e)
        {
            WordApp.Selection.ClearFormatting();
            WordApp.Selection.set_Style("正文");
        }

        //设置主标题格式
        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            WordApp.Selection.set_Style("合同主标题");
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

        //创建合同主标题字体样式
        private void create_style_biaoti()
        {
            Style style = null;
            //添加和设置正文格式
            try
            {
                style = WordApp.ActiveDocument.Styles["合同主标题"];
            }
            catch (Exception)
            {
                WordApp.ActiveDocument.Styles.Add("合同主标题");
                style = WordApp.ActiveDocument.Styles["合同主标题"];
            }

            style.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpace1pt5;  //行距为1.5倍行距
            style.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;       //文字居中
            style.set_BaseStyle("");                                               //设置基准样式为（无样式）
            style.set_NextParagraphStyle("正文");                                   //设置后续样式为“正文”
            style.ParagraphFormat.FirstLineIndent = 0;                           //首行缩进0个字符

            style.Font.NameFarEast = "方正小标宋简体";                           //中文字体
            style.Font.NameAscii = "仿宋";                                       //英文格式
            style.Font.NameOther = "仿宋";                                       //字符格式
            style.Font.Name = "方正小标宋简体";                                  //格式名称
            style.Font.Size = 40;                                                //设置字体大小
        }

        //创建合同编号部分的字体样式
        private void create_style_bianhao()
        {

        }
 
    }
}
