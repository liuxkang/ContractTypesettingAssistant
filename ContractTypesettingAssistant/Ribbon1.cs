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
            button2.Enabled = false;
            button3.Enabled = false;
            button4.Enabled = false;
            button5.Enabled = false;
            button6.Enabled = false;
            button7.Enabled = false;
            button8.Enabled = false;
        }

        //设置合同模板全局，增加上下左右边距等
        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            //创建各种样式
            create_style_zhengwen();            //插件启动时创建 正文 格式
            create_style_biaoti();              //插件启动时创建 合同标题 格式
            create_style_bianhao();             //插件启动时创建 编号部分 格式
            create_style_zhangjie();            //插件启动时创建 合同章节 格式
            create_style_table();               //插件启动时创建 合同表格 格式
            create_style_yemei();               //插件启动时创建 合同页眉 格式

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
            button2.Enabled = true;
            button3.Enabled = true;
            button4.Enabled = true;
            button5.Enabled = true;
            button6.Enabled = true;
            button7.Enabled = true;
            button8.Enabled = true;
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
            WordApp.Selection.ClearFormatting();
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
            style.ParagraphFormat.LeftIndent = WordApp.CentimetersToPoints(0.0f); ;           //左边缩进(以磅为单位)
            style.ParagraphFormat.RightIndent = WordApp.CentimetersToPoints(0.0f); ;         //右边缩进(以磅为单位)
            style.ParagraphFormat.SpaceBefore = 0f;       //段前间距（以磅为单位）
            style.ParagraphFormat.SpaceBeforeAuto = 0;                     //自动段前间距
            style.ParagraphFormat.SpaceAfter = 0f;          //段后间距
            style.ParagraphFormat.SpaceAfterAuto = 0;                       //自动段后间距
            style.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;                  //行距设置：固定值
            style.ParagraphFormat.LineSpacing = 24;                                                                             //行距设置：24磅
            style.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;     //对齐方式：左对齐
            style.ParagraphFormat.WidowControl = 0;             //孤行控制：不勾选
            style.ParagraphFormat.KeepWithNext = 0;             //与下段同页：不勾选
            style.ParagraphFormat.KeepTogether = 0;             //段中不分页：不勾选
            style.ParagraphFormat.PageBreakBefore = 0;         //段前分页：不勾选
            style.ParagraphFormat.NoLineNumber = 0;             //取消行号：不勾选
            style.ParagraphFormat.Hyphenation = 0;                //取消断字：不勾选
            style.ParagraphFormat.FirstLineIndent = WordApp.CentimetersToPoints(0.0f);         //首行缩进或悬挂(以磅为单位)：2字符
            style.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevelBodyText;     //大纲级别：正文文本
            style.ParagraphFormat.CharacterUnitLeftIndent = 0;                          //左边缩进(以字符为单位)
            style.ParagraphFormat.CharacterUnitRightIndent = 0;                         //右边缩进(以字符为单位)
            style.ParagraphFormat.CharacterUnitFirstLineIndent = 2;                     //首行或悬挂（以字符为单位）
            style.ParagraphFormat.LineUnitBefore = 0;                                         //段前间距（行）
            style.ParagraphFormat.LineUnitAfter = 0;                                            //段后间距（行）
            style.ParagraphFormat.MirrorIndents = 0;                                            //在相同样式的段落间不添加空格
            style.ParagraphFormat.TextboxTightWrap = WdTextboxTightWrap.wdTightNone;     //未知
            style.ParagraphFormat.AutoAdjustRightIndent = 0;                            //自动调整右间距
            style.ParagraphFormat.DisableLineHeightGrid = 0;                                //段落中的字符与行网格不进行对齐
            style.ParagraphFormat.FarEastLineBreakControl = 0;                              //设置指定文档的换行控制级别。
            style.ParagraphFormat.WordWrap = 0;         //在指定段落或文本框的西文单词中间断字换行
            style.ParagraphFormat.HangingPunctuation = 0;       //允许标点溢出边界
            style.ParagraphFormat.AddSpaceBetweenFarEastAndAlpha = 0;       //中文文字和拉丁文字之间自动添加空格
            style.ParagraphFormat.AddSpaceBetweenFarEastAndDigit = 0;          //中文文字和数字之间添加空格
            style.ParagraphFormat.BaseLineAlignment = WdBaselineAlignment.wdBaselineAlignAuto; //整活动文档中的基线字体对齐方式：自动

            style.Font.NameFarEast = "仿宋";                                                //中文字体
            style.Font.NameAscii = "+西文正文";                                                  //英文格式
            style.Font.NameOther = "仿宋";                                                  //字符格式
            style.Font.Name = "仿宋";                                                       //格式名称
            style.Font.Size = 14;                                                           //正文字号"四号"
            style.Font.Bold = 0;

        }

        //创建合同章节样式
        private void create_style_zhangjie()
        {
            Style style = null;
            //添加和设置正文格式
            try
            {
                style = WordApp.ActiveDocument.Styles["合同章节"];
            }
            catch (Exception)
            {
                WordApp.ActiveDocument.Styles.Add("合同章节");
                style = WordApp.ActiveDocument.Styles["合同章节"];
            }
            style.set_BaseStyle("正文");
            style.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel2;
            style.set_NextParagraphStyle("正文");
            style.Font.Size = 14;
            style.Font.Bold = 1;

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

            style.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpace1pt5;          //行距为1.5倍行距
            style.ParagraphFormat.CharacterUnitLeftIndent = 0;                              //左边缩进0个字符
            style.ParagraphFormat.LeftIndent = 0;                                           //左边缩进0厘米
            style.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;  //文字居中
            style.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel1;            //设置大纲级别1级
            style.set_BaseStyle("");                                                        //设置基准样式为（无样式）
            style.set_NextParagraphStyle("正文");                                           //设置后续样式为“正文”
            style.ParagraphFormat.FirstLineIndent = 0;                                      //首行缩进0个字符
            style.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;  //居中对齐

            style.Font.NameFarEast = "方正小标宋简体";                                      //中文字体
            style.Font.NameAscii = "仿宋";                                                  //英文格式
            style.Font.NameOther = "仿宋";                                                  //字符格式
            style.Font.Name = "方正小标宋简体";                                             //格式名称
            style.Font.Size = 40;                                                           //设置字体大小
            style.Font.Bold = 0;
        }

        //创建合同编号部分的字体样式
        private void create_style_bianhao()
        {
            Style style = null;
            try
            {
                style = WordApp.ActiveDocument.Styles["编号部分"];
            }
            catch(Exception)
            {
                WordApp.ActiveDocument.Styles.Add("编号部分");
                style = WordApp.ActiveDocument.Styles["编号部分"];
            }
            style.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;       //行距设置为固定值
            style.ParagraphFormat.LineSpacing = 30;                                         //行距30磅
            style.ParagraphFormat.CharacterUnitLeftIndent = 0;                              //左边缩进0个字符
            style.ParagraphFormat.LeftIndent = 0;                                           //左边缩进0厘米
            style.set_BaseStyle("");                                                        //设置基准样式为（无样式）
            style.set_NextParagraphStyle("正文");                                           //设置后续样式为“正文”
            style.ParagraphFormat.FirstLineIndent = 0;                                      //首行缩进0个字符
            style.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;    //左对齐

            style.Font.NameFarEast = "仿宋";                                                //中文字体
            style.Font.NameAscii = "仿宋";                                                  //英文格式
            style.Font.NameOther = "仿宋";                                                  //字符格式
            style.Font.Name = "仿宋";                                                       //格式名称
            style.Font.Size = 17;                                                           //设置字体大小
            style.Font.Bold = 1;                                                            //设置为粗体

        }

        //设置编号部分字体
        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            WordApp.Selection.ClearFormatting();
            WordApp.Selection.set_Style("编号部分");
        }

        //首页页底签订时间 签订地点部分
        private void button4_Click(object sender, RibbonControlEventArgs e)
        {
            WordApp.Selection.ClearFormatting();
            WordApp.Selection.TypeText("  签订时间：201 年   月   日");
            WordApp.Selection.set_Style("正文");
            WordApp.Selection.TypeParagraph();

            WordApp.Selection.TypeText("  签订地点：云南.河口");
            WordApp.Selection.set_Style("正文");
            WordApp.Selection.InsertBreak(WdBreakType.wdPageBreak);

        }

        //创建合同表格内字体样式
        private void create_style_table()
        {
            Style style = null;
            try
            {
                style = WordApp.ActiveDocument.Styles["合同表格"];
            }
            catch (Exception)
            {
                WordApp.ActiveDocument.Styles.Add("合同表格");
                style = WordApp.ActiveDocument.Styles["合同表格"];
            }
            style.set_BaseStyle("");
            style.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;              //单倍行距
            style.ParagraphFormat.CharacterUnitLeftIndent = 0;                                    //左边缩进0个字符
            style.ParagraphFormat.LeftIndent = 0;                                                 //左边缩进0厘米
            style.ParagraphFormat.FirstLineIndent = 0;                                            //首行缩进0个字符
            style.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevelBodyText;           //大纲级别正文文本
            style.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;          //左对齐

            style.Font.NameFarEast = "仿宋";                                                      //中文字体
            style.Font.NameAscii = "仿宋";                                                        //英文格式
            style.Font.NameOther = "仿宋";                                                        //字符格式
            style.Font.Name = "仿宋";                                                             //格式名称
            style.Font.Size = 12;                                                                 //设置字体大小 小四号
            style.Font.Bold = 0;                                                                  //不加粗

        }

        //创建合同页眉
        private void create_style_yemei()
        {
            Style style = null;
            try
            {
                style = WordApp.ActiveDocument.Styles["合同页眉"];
            }
            catch (Exception)
            {
                WordApp.ActiveDocument.Styles.Add("合同页眉");
                style = WordApp.ActiveDocument.Styles["合同页眉"];
            }
            style.set_BaseStyle("");
            style.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;               //行距为固定值
            style.ParagraphFormat.LineSpacing = 12;                                                 //行距为12磅
            style.ParagraphFormat.CharacterUnitLeftIndent = 0;                                      //左边缩进0个字符
            style.ParagraphFormat.LeftIndent = 0;                                                   //左边缩进0厘米
            style.set_NextParagraphStyle("合同页眉");
            style.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;            //左对齐
            style.ParagraphFormat.FirstLineIndent = 0;
            style.ParagraphFormat.CharacterUnitFirstLineIndent = 0;                                 //首行缩进两个字符

            style.Font.NameFarEast = "仿宋";                                                      //中文字体
            style.Font.NameAscii = "仿宋";                                                        //英文格式
            style.Font.NameOther = "仿宋";                                                        //字符格式
            style.Font.Name = "仿宋";                                                             //格式名称
            style.Font.Size = 9;                                                                 //设置字体大小 小五号
            style.Font.Bold = 0;

        }

        //合同章节部分
        private void button6_Click(object sender, RibbonControlEventArgs e)
        {
            WordApp.Selection.ClearFormatting();
            WordApp.Selection.set_Style("合同章节");
        }

        //合同表格
        private void button7_Click(object sender, RibbonControlEventArgs e)
        {
            WordApp.Selection.ClearFormatting();
            WordApp.Selection.set_Style("合同表格");
        }

        //合同页眉
        private void button8_Click(object sender, RibbonControlEventArgs e)
        {
            WordApp.ActiveWindow.View.SeekView = WdSeekView.wdSeekCurrentPageHeader;
            WordApp.Selection.HeaderFooter.LinkToPrevious = true;
            WordApp.Selection.HeaderFooter.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            WordApp.Selection.set_Style("合同页眉");
            WordApp.ActiveWindow.View.SeekView = WdSeekView.wdSeekMainDocument;
        }
    }
}
