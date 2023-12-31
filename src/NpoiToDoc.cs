﻿using DBToDocUtils.Models;
using NPOI.XWPF.UserModel;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace DBToDocUtils
{
    public class NpoiToDoc
    {

        IBaseService service = new BaseService();
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="type"></param>
        public NpoiToDoc(int type)
        {
            GetDBService(type);
        }
        /// <summary>
        /// 获取数据库方法
        /// </summary>
        /// <param name="type">数据库类型：0 SQLServer、1 MySQL 、2 SQlite 、 pg 3  默认是SQLServer</param>
        private void GetDBService(int type)
        {
            if (type == 0)
            {
                service = new BaseService();
            }
            else if (type == 1)
            {
                service = new BaseServiceMysql();
            }
            else if (type == 2)
            {
                // 获取文件名            
                service = new BaseServiceSqlite();
            }
            else if (type == 3)
            {
                service = new BaseServicePgsql();
            }
            else
            {
                service = new BaseService();
            }
        }

        /// <summary>
        /// 生成word文档
        /// </summary>
        /// <param name="conStr"></param>
        /// <param name="db"></param>    
        /// <param name="path"></param> 
        /// <param name="checkStr">表,视图,存储过程 支持多选</param>
        public void CreateToWord(string conStr, string db, string path, string checkStr = "表,视图,存储过程")
        {
            List<string> checkList = new List<string>();
            checkList = checkStr.Split(',').ToList();
            List<TableModel> list = service.GetDBTableList(conStr, db);
            XWPFDocument doc = new XWPFDocument(); //创建新的word文档

            XWPFParagraph p1 = doc.CreateParagraph();//向新文档中添加段落

            p1.Alignment = ParagraphAlignment.CENTER;
            XWPFRun r1 = p1.CreateRun();
            r1.FontFamily = "微软雅黑";
            r1.FontSize = 28;
            r1.IsBold = true;
            //向该段落中添加文字
            r1.SetText(db + "数据库说明文档");
            r1.AddCarriageReturn();
            #region 创建一个表格
            if (checkList.Where(m => m.Equals("表")).Count() > 0)
            {
                XWPFParagraph tableHead = doc.CreateParagraph();
                tableHead.Alignment = ParagraphAlignment.LEFT;
                XWPFRun tableHeadR = tableHead.CreateRun();
                tableHeadR.FontFamily = "微软雅黑";
                tableHeadR.FontSize = 22;
                tableHeadR.IsBold = true;
                tableHeadR.SetText("一、表结构");
                tableHeadR.AddCarriageReturn();

                if (list.Count > 0)
                {
                    foreach (var item in list)
                    {

                        XWPFParagraph p3 = doc.CreateParagraph();   //向新文档中添加段落
                        p3.Alignment = ParagraphAlignment.LEFT;
                        XWPFRun r3 = p3.CreateRun();                //向该段落中添加文字
                        r3.FontFamily = "微软雅黑";
                        r3.FontSize = 18;
                        r3.IsBold = true;
                        if (item.tableDesc != null && item.tableDesc != "")
                        {
                            r3.SetText("表名:" + item.tableName + "(" + item.tableDesc + ")");
                        }
                        else
                        {
                            r3.SetText("表名:" + item.tableName);
                        }

                        //从第二行开始 因为第一行是表头
                        int i = 1;
                        var tabledetaillist = service.GetTableDetail(item.tableName, conStr, db);
                        XWPFTable table = doc.CreateTable(tabledetaillist.Count + 1, 9);
                        table.Width = 5000;

                        #region 设置表头               

                        //table.GetRow(0).GetCell(0).SetText("数据库名称");
                        XWPFParagraph pI = table.GetRow(0).GetCell(0).AddParagraph();
                        XWPFRun rI = pI.CreateRun();
                        rI.FontFamily = "微软雅黑";
                        rI.FontSize = 10;
                        rI.IsBold = true;
                        rI.SetText("序号");


                        XWPFParagraph pI1 = table.GetRow(0).GetCell(1).AddParagraph();
                        XWPFRun rI1 = pI1.CreateRun();
                        rI1.FontFamily = "微软雅黑";
                        rI1.FontSize = 10;
                        rI1.IsBold = true;
                        rI1.SetText("字段名称");

                        XWPFParagraph pI2 = table.GetRow(0).GetCell(2).AddParagraph();
                        XWPFRun rI2 = pI2.CreateRun();
                        rI2.FontFamily = "微软雅黑";
                        rI2.FontSize = 10;
                        rI2.IsBold = true;
                        rI2.SetText("标识");

                        XWPFParagraph pI3 = table.GetRow(0).GetCell(3).AddParagraph();
                        XWPFRun rI3 = pI3.CreateRun();
                        rI3.FontFamily = "微软雅黑";
                        rI3.FontSize = 10;
                        rI3.IsBold = true;
                        rI3.SetText("主键");

                        XWPFParagraph pI4 = table.GetRow(0).GetCell(4).AddParagraph();
                        XWPFRun rI4 = pI4.CreateRun();
                        rI4.FontFamily = "微软雅黑";
                        rI4.FontSize = 10;
                        rI4.IsBold = true;
                        rI4.SetText("字段类型");

                        XWPFParagraph pI5 = table.GetRow(0).GetCell(5).AddParagraph();
                        XWPFRun rI5 = pI5.CreateRun();
                        rI5.FontFamily = "微软雅黑";
                        rI5.FontSize = 10;
                        rI5.IsBold = true;
                        rI5.SetText("字段长度");

                        XWPFParagraph pI6 = table.GetRow(0).GetCell(6).AddParagraph();
                        XWPFRun rI6 = pI6.CreateRun();
                        rI6.FontFamily = "微软雅黑";
                        rI6.FontSize = 10;
                        rI6.IsBold = true;
                        rI6.SetText("允许空");


                        XWPFParagraph pI7 = table.GetRow(0).GetCell(7).AddParagraph();
                        XWPFRun rI7 = pI7.CreateRun();
                        rI7.FontFamily = "微软雅黑";
                        rI7.FontSize = 10;
                        rI7.IsBold = true;
                        rI7.SetText("字段默认值");

                        XWPFParagraph pI8 = table.GetRow(0).GetCell(8).AddParagraph();
                        XWPFRun rI8 = pI8.CreateRun();
                        rI8.FontFamily = "微软雅黑";
                        rI8.FontSize = 10;
                        rI8.IsBold = true;
                        rI8.SetText("字段说明");

                        #endregion
                        if (tabledetaillist != null && tabledetaillist.Count > 0)
                        {
                            foreach (var itm in tabledetaillist)
                            {
                                //第一列
                                XWPFParagraph pIO = table.GetRow(i).GetCell(0).AddParagraph();
                                XWPFRun rIO = pIO.CreateRun();
                                //rIO.FontFamily = "微软雅黑";
                                rIO.FontSize = 10;
                                rIO.IsBold = false;
                                rIO.SetText(itm.index.ToString());

                                //第二列
                                XWPFParagraph pIO2 = table.GetRow(i).GetCell(1).AddParagraph();
                                XWPFRun rIO2 = pIO2.CreateRun();
                                //rIO2.FontFamily = "微软雅黑";
                                rIO2.FontSize = 10;
                                rIO2.IsBold = false;
                                rIO2.SetText(itm.Title);


                                XWPFParagraph pIO3 = table.GetRow(i).GetCell(2).AddParagraph();
                                XWPFRun rIO3 = pIO3.CreateRun();

                                rIO3.FontSize = 10;
                                rIO3.IsBold = false;
                                rIO3.SetText(itm.isMark.ToString());

                                XWPFParagraph pIO4 = table.GetRow(i).GetCell(3).AddParagraph();
                                XWPFRun rIO4 = pIO4.CreateRun();
                                rIO4.FontSize = 10;
                                rIO4.IsBold = false;
                                rIO4.SetText(itm.isPK.ToString());

                                XWPFParagraph pIO5 = table.GetRow(i).GetCell(4).AddParagraph();
                                XWPFRun rIO5 = pIO5.CreateRun();
                                rIO5.FontSize = 10;
                                rIO5.IsBold = false;
                                rIO5.SetText(itm.FieldType);

                                XWPFParagraph pIO6 = table.GetRow(i).GetCell(5).AddParagraph();
                                XWPFRun rIO6 = pIO6.CreateRun();
                                rIO6.FontSize = 10;
                                rIO6.IsBold = false;
                                rIO6.SetText(itm.fieldLenth.ToString());

                                XWPFParagraph pIO7 = table.GetRow(i).GetCell(6).AddParagraph();
                                XWPFRun rIO7 = pIO7.CreateRun();
                                rIO7.FontSize = 10;
                                rIO7.IsBold = false;
                                rIO7.SetText(itm.isAllowEmpty.ToString());

                                XWPFParagraph pIO8 = table.GetRow(i).GetCell(7).AddParagraph();
                                XWPFRun rIO8 = pIO8.CreateRun();
                                rIO8.FontSize = 10;
                                rIO8.IsBold = false;
                                rIO8.SetText(itm.defaultValue.ToString());

                                XWPFParagraph pIO9 = table.GetRow(i).GetCell(8).AddParagraph();
                                XWPFRun rIO9 = pIO9.CreateRun();
                                //rIO9.FontFamily = "微软雅黑";
                                rIO9.FontSize = 10;
                                rIO9.IsBold = false;
                                rIO9.SetText(itm.fieldDesc);

                                i++;
                            }
                        }

                        XWPFParagraph psql3 = doc.CreateParagraph();   //向新文档中添加段落
                        psql3.Alignment = ParagraphAlignment.LEFT;
                        XWPFRun rsql3 = psql3.CreateRun();                //向该段落中添加文字
                        rsql3.FontFamily = "微软雅黑";
                        rsql3.FontSize = 10;
                        rsql3.IsBold = true;
                        rsql3.SetText(item.tableName + "建表脚本");

                        XWPFParagraph psql4 = doc.CreateParagraph();   //向新文档中添加段落
                        psql4.Alignment = ParagraphAlignment.LEFT;
                        XWPFRun rsql4 = psql4.CreateRun();                //向该段落中添加文字
                        rsql4.FontFamily = "微软雅黑";
                        rsql4.FontSize = 10;
                        rsql4.IsBold = false;
                        rsql4.SetText(service.GetTableSQL(item.tableName, conStr));

                    }
                }
            }
            #endregion

            #region 存储过程
            if (checkList.Where(m => m.Equals("存储过程")).Count() > 0)
            {
                XWPFParagraph p2 = doc.CreateParagraph();
                XWPFRun r2 = p2.CreateRun();
                r2.FontFamily = "微软雅黑";
                r2.IsBold = true;
                r2.FontSize = 22;
                r2.SetText("二、存储过程");
                r2.AddCarriageReturn();
                List<ProcModel> proclist = new List<ProcModel>();
                proclist = service.GetProcList(conStr, db);
                if (proclist != null && proclist.Count > 0)
                {
                    foreach (var item in proclist)
                    {
                        //存储过程名称
                        XWPFParagraph pro1 = doc.CreateParagraph();
                        XWPFRun rpro1 = pro1.CreateRun();
                        rpro1.FontSize = 14;
                        rpro1.IsBold = true;
                        rpro1.SetText("存储过程名称：" + item.procName);
                        //存储过程 详情
                        XWPFParagraph pro2 = doc.CreateParagraph();
                        XWPFRun rpro2 = pro2.CreateRun();
                        rpro2.FontSize = 12;
                        rpro2.SetText(string.IsNullOrWhiteSpace(item.proDerails) ? "" : item.proDerails);
                    }
                }
            }
            #endregion

            #region 视图
            XWPFParagraph v2 = doc.CreateParagraph();
            XWPFRun vr2 = v2.CreateRun();
            vr2.IsBold = true;
            vr2.FontFamily = "微软雅黑";
            vr2.FontSize = 22;
            vr2.SetText("三、视图");
            vr2.AddCarriageReturn();
            if (checkList.Where(m => m.Equals("视图")).Count() > 0)
            {

                List<ViewModel> viewlist = new List<ViewModel>();
                viewlist = service.GetViewList(conStr, db);
                if (viewlist.Count > 0)
                {
                    foreach (var item in viewlist)
                    {
                        //存储过程名称
                        XWPFParagraph vro1 = doc.CreateParagraph();
                        XWPFRun vpro1 = vro1.CreateRun();
                        vpro1.FontSize = 14;
                        vpro1.IsBold = true;
                        vpro1.SetText("视图名称：" + item.viewName);
                        //存储过程 详情
                        XWPFParagraph vro2 = doc.CreateParagraph();
                        XWPFRun vpro2 = vro2.CreateRun();
                        vpro2.FontSize = 12;
                        vpro2.SetText(string.IsNullOrWhiteSpace(item.viewDerails) ? "" : item.viewDerails);
                    }
                }
            }
            #endregion

            string filename = path +"\\"+ db + "-数据库说明文档.doc";//文件名
            /* SaveFileDialog saveDialog = new SaveFileDialog();
             saveDialog.DefaultExt = "doc";
             saveDialog.Filter = "Word文件|*.doc";
             saveDialog.FileName = filename;
             saveDialog.ShowDialog();
             filename = saveDialog.FileName;
             if (filename.IndexOf(":") < 0) return; //被点了取消    */
            FileStream sw1 = new FileStream(filename, FileMode.Create);
            doc.Write(sw1);
            sw1.Close();
            //System.Diagnostics.Process.Start(filename);

        }

        /// <summary>
        /// 生成html文件
        /// </summary>
        /// <param name="conStr"></param>
        /// <param name="db"></param>
        /// <param name="path"></param>
        /// <param name="checkStr">表,视图,存储过程 支持多选</param>
        public void CreateToHtml(string conStr, string db, string path, string checkStr = "表,视图,存储过程")
        {

            List<string> checkList = new List<string>();
            checkList = checkStr.Split(',').ToList();
            StringBuilder sb = new StringBuilder();
            sb.Append("<html><meta charset=\"utf-8\" /><meta http-equiv = \"Content-Language\" content = \"zh-CN\" >");
            sb.Append("<head><title>数据库说明文档</title><body>");
            sb.Append("<style type=\"text/css\">\n");
            sb.Append("body { font-size: 9pt;}\n");
            sb.Append(".styledb { font-size: 14px; }\n");
            sb.Append(".styletab {font-size: 14px;padding-top: 15px; }\n</style></head><body>");
            sb.Append("<h1 style=\"text-align:center;\">" + db + "数据库说明文档</h1>");

            List<TableModel> list = service.GetDBTableList(conStr, db);

            #region 创建一个表格
            if (checkList.Where(m => m.Equals("表")).Count() > 0)
            {
                sb.Append("<h2>一、表结构</h2>");
                sb.Append("");
                sb.Append("");
                if (list.Count > 0)
                {
                    foreach (var item in list)
                    {
                        if (item.tableDesc != null && item.tableDesc != "")
                        {
                            sb.Append("<h3>表名:" + item.tableName + "(" + item.tableDesc + ")</h3>");
                        }
                        else
                        {
                            sb.Append("<h3>表名:" + item.tableName + "</h3>");
                        }
                        sb.Append(" <table cellspacing=\"0\" cellpadding=\"5\" border=\"1\" width=\"100%\" bordercolorlight=\"#4F7FC9\" bordercolordark=\"#D3D8E0\">");
                        sb.Append("<thead bgcolor=\"#E3EFFF\"> <th>序号</th><th>字段名称</th><th>标识</th><th>主键</th><th>字段类型</th><th>字段长度</th><th>允许空值</th><th>字段默认值</th><th>字段备注</th></thead>");
                        sb.Append("<tbody>");
                        //从第二行开始 因为第一行是表头
                        int i = 1;
                        var tabledetaillist = service.GetTableDetail(item.tableName, conStr, db);


                        if (tabledetaillist != null && tabledetaillist.Count > 0)
                        {
                            foreach (var itm in tabledetaillist)
                            {
                                sb.Append("<tr>");
                                sb.Append("<td>" + itm.index + "</td>");
                                sb.Append("<td>" + itm.Title + "</td>");
                                sb.Append("<td>" + itm.isMark + "</td>");
                                sb.Append("<td>" + itm.isPK + "</td>");
                                sb.Append("<td>" + itm.FieldType + "</td>");
                                sb.Append("<td>" + itm.fieldLenth + "</td>");
                                sb.Append("<td>" + itm.isAllowEmpty + "</td>");
                                sb.Append("<td>" + itm.defaultValue + "</td>");
                                sb.Append("<td>" + itm.fieldDesc + "</td>");
                                sb.Append("</tr>");
                                i++;
                            }
                        }
                        sb.Append("</tbody></table>");

                        sb.Append("<h4>" + item.tableName + "建表脚本</h4><br/>");
                        sb.Append("<span>" + service.GetTableSQL(item.tableName, conStr) + "</span>");


                    }
                }
            }
            #endregion

            #region 存储过程
            if (checkList.Where(m => m.Equals("存储过程")).Count() > 0)
            {
                List<ProcModel> proclist = new List<ProcModel>();
                proclist = service.GetProcList(conStr, db);
                sb.Append("<h2>二、存储过程</h2>");
                if (proclist != null && proclist.Count > 0)
                {
                    foreach (var item in proclist)
                    {
                        sb.Append("<h3>存储过程名称：" + item.procName + "</h3>");
                        sb.Append("<span>" + item.proDerails + "</span>");
                    }
                }
            }
            #endregion

            #region 视图
            if (checkList.Where(m => m.Equals("视图")).Count() > 0)
            {
                List<ViewModel> viewlist = new List<ViewModel>();
                viewlist = service.GetViewList(conStr, db);
                sb.Append("<h2>三、视图</h2>");
                if (viewlist.Count > 0)
                {

                    foreach (var item in viewlist)
                    {
                        sb.Append("<h3>视图名称：" + item.viewName + "</h3>");
                        sb.Append("<span>" + item.viewDerails + "</span>");
                    }
                }
            }
            #endregion

            sb.Append("</body></html>");
            sb.ToString();
            string filename = path + "\\" + db + "-数据库说明文档.html";//文件名
         /*   SaveFileDialog saveDialog = new SaveFileDialog();
            saveDialog.DefaultExt = "html";
            saveDialog.Filter = "html文件|*.html";
            saveDialog.FileName = filename;
            saveDialog.ShowDialog();
            filename = saveDialog.FileName;
            if (filename.IndexOf(":") < 0) return; //被点了取消    */     
            StreamWriter sw1 = new StreamWriter(filename, false);
            sw1.WriteLine(sb);
            sw1.Close();
            //System.Diagnostics.Process.Start(filename);

        }


        /// <summary>
        /// 生成markwdown文件
        /// </summary>
        /// <param name="conStr"></param>
        /// <param name="db"></param>
        /// <param name="path"></param>
        /// <param name="checkStr"></param>
        public void CreateToMarkDown(string conStr, string db,string path, string checkStr = "表,视图,存储过程")
        {
            List<string> checkList = new List<string>();
            checkList = checkStr.Split(',').ToList();
            StringBuilder sb = new StringBuilder();
            sb.Append(" # " + db + "数据库说明文档\n\n\n");
            List<TableModel> list = service.GetDBTableList(conStr, db);

            #region 创建一个表格
            if (checkList.Where(m => m.Equals("表")).Count() > 0)
            {

                sb.Append("## 一、表结构\n");

                if (list.Count > 0)
                {
                    foreach (var item in list)
                    {
                        if (item.tableDesc != null && item.tableDesc != "")
                        {

                            sb.Append("### 表名:" + item.tableName + "(" + item.tableDesc + ")\n");
                        }
                        else
                        {
                            sb.Append("### 表名:" + item.tableName + "\n");
                        }

                        sb.Append("| 序号 | 字段名称 | 标识 | 主键 | 字段类型 | 字段长度 | 允许空值 | 字段默认值 | 字段备注 |\n");
                        sb.Append("| :------: | :------: | :------: | :------: | :------: | :------: | :------: | :------: | :------: |\n");
                        //从第二行开始 因为第一行是表头
                        int i = 1;
                        var tabledetaillist = service.GetTableDetail(item.tableName, conStr, db);


                        if (tabledetaillist != null && tabledetaillist.Count > 0)
                        {
                            foreach (var itm in tabledetaillist)
                            {
                                sb.Append("| " + itm.index + " | " + itm.Title + " | " + itm.isMark + " | " + itm.isPK + " | " + itm.FieldType + " | " + itm.fieldLenth + " | " + itm.isAllowEmpty + "  | " + itm.defaultValue + " | " + itm.fieldDesc + " | \n");
                                i++;
                            }
                        }

                        sb.Append("#### " + item.tableName + "建表脚本\n");
                        sb.Append("```sql \n " + service.GetTableSQL(item.tableName, conStr) + "\n");
                        sb.Append("```\n");

                    }
                }
            }
            #endregion

            #region 存储过程
            if (checkList.Where(m => m.Equals("存储过程")).Count() > 0)
            {
                List<ProcModel> proclist = new List<ProcModel>();
                proclist = service.GetProcList(conStr, db);
                sb.Append("## 二、存储过程\n");
                if (proclist != null && proclist.Count > 0)
                {
                    // sb.Append("<h2>二、存储过程</h2>");

                    foreach (var item in proclist)
                    {
                        sb.Append("### 存储过程名称：" + item.procName + "\n");
                        sb.Append("```sql \n" + item.proDerails + "\n");
                        sb.Append("```\n");
                    }
                }
                else
                {
                    sb.Append("无\n");
                }
            }
            #endregion

            #region 视图
            if (checkList.Where(m => m.Equals("视图")).Count() > 0)
            {
                List<ViewModel> viewlist = new List<ViewModel>();
                viewlist = service.GetViewList(conStr, db);
                sb.Append("## 三、视图\n");
                if (viewlist.Count > 0)
                {
                    foreach (var item in viewlist)
                    {
                        sb.Append("### 视图名称：" + item.viewName + "\n");
                        sb.Append("```sql \n" + item.viewDerails + "\n");
                        sb.Append("```\n");
                    }
                }
                else
                {
                    sb.Append("无\n");
                }
            }
            #endregion

            sb.ToString();
            Html2Markdown.Converter converter = new Html2Markdown.Converter();
            string mdStr = converter.Convert(sb.ToString());          
            string filename = path + "\\" + db + "-数据库说明文档.md";//文件名
          /*  SaveFileDialog saveDialog = new SaveFileDialog();
            saveDialog.DefaultExt = "md";
            saveDialog.Filter = "md文件|*.md";
            saveDialog.FileName = filename;
            saveDialog.ShowDialog();
            filename = saveDialog.FileName;
            if (filename.IndexOf(":") < 0) return; //被点了取消 */        
            StreamWriter sw1 = new StreamWriter(filename, false);
            sw1.WriteLine(mdStr);
            sw1.Close();
            //System.Diagnostics.Process.Start(filename);
        }
    }
}
