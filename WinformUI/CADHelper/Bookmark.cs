using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Spire.Doc;
using Spire.Doc.Fields;
using Spire.Doc.Interface;
using Spire.Doc.Documents;

namespace DotNetARX
{
    public class Bookmark
    {
        private Document doc = null;
        public Bookmark(string filePath)
        {
            doc = new Document();
            doc.LoadFromFile(filePath);
        }


        /// <summary>
        /// 用文本替换指定书签的内容
        /// </summary>
        /// <param name="bookmarkName">书签名</param>
        /// <param name="text">文本</param>
        /// <param name="saveFormatting">删除原始书签内容时，是否保留格式</param>
        /// <returns>TextRange</returns>
        public TextRange ReplaceContent(string bookmarkName, string text, bool saveFormatting)
        {
            BookmarksNavigator navigator = new BookmarksNavigator(doc);
            navigator.MoveToBookmark(bookmarkName);//指向特定书签           
            navigator.DeleteBookmarkContent(saveFormatting);//删除原有书签内容     
            Spire.Doc.Interface.ITextRange textRange = navigator.InsertText(text);//写入文本
            return textRange as TextRange;
        }
        public string GetBookmarkContent(string bookmarkName)
        {
            BookmarksNavigator navigator = new BookmarksNavigator(doc);
            navigator.MoveToBookmark(bookmarkName);//指向指定书签
            TextBodyPart textBodyPart = navigator.GetBookmarkContent();
            string text = null;
            foreach (var item in textBodyPart.BodyItems)
            {
                if (item is Paragraph)
                {
                    foreach (var childObject in (item as Paragraph).ChildObjects)
                    {
                        if (childObject is TextRange)
                        {
                            text += (childObject as TextRange).Text;
                        }
                    }
                }
            }
            return text;

        }
        public void Save(string outputPath)
        {
            //TODO 2021.11.15  1、标准2000机柜生成过程中报错！
            doc.SaveToFile(outputPath);
        }
        /// <summary>
        /// 用表格替换指定书签的内容
        /// </summary>
        /// <param name="bookmarkName">书签名</param>
        /// <param name="table">Table实例</param>
        public void ReplaceContent(string bookmarkName, Table table)
        {
            BookmarksNavigator navigator = new BookmarksNavigator(doc);
            navigator.MoveToBookmark(bookmarkName);
            TextBodyPart body = new TextBodyPart(doc);
            body.BodyItems.Add(table);
            navigator.ReplaceBookmarkContent(body);
        }

        /// <summary>
        /// 创建表格并写入数据，返回Table对象
        /// </summary>
        /// <param name="rowsNum">行数</param>
        /// <param name="columnsNum">列数</param>
        /// <param name="columnWidth">列宽</param>
        /// <param name="horizontalAlignment">水平对齐方式</param>
        /// <param name="datatable">DataTable实例</param>
        /// <returns></returns>
        public Table CreateTable(int rowsNum, int columnsNum, float columnWidth, RowAlignment horizontalAlignment, System.Data.DataTable datatable)
        {
            Table table = new Table(doc, true, 1f);//初始化Table对象
            table.ResetCells(rowsNum, columnsNum);//设置行数和列数
            //填充数据
            for (int i = 0; i < datatable.Rows.Count; i++)
            {
                for (int j = 0; j < datatable.Columns.Count; j++)
                {
                    table.Rows[i].Cells[j].AddParagraph().AppendText(datatable.Rows[i][j].ToString());
                }
            }
            //设置列宽
            for (int i = 0; i < rowsNum; i++)
            {
                for (int j = 0; j < columnsNum; j++)
                {
                    table.Rows[i].Cells[j].Width = columnWidth;
                }
            }
            table.TableFormat.HorizontalAlignment = horizontalAlignment;//表格水平对齐方式
            return table;
        }
        /// <summary>
        /// 设置表格的值
        /// </summary>
        /// <param name="rowIndex"></param>
        /// <param name="colIndex"></param>
        /// <param name="value"></param>
        public void SetCellValue(int rowIndex, int colIndex, string value)
        {
            rowIndex += 6;
            Section section = doc.Sections[0];
            Table table = section.Tables[0] as Table;
            TableCell cell = table.Rows[rowIndex].Cells[colIndex];
            cell.Paragraphs[0].Text = value;
        }
        public void SetCellValue(int setctionIndex, int rowIndex, int colIndex, string value)
        {
            rowIndex += 1;
            Section section = doc.Sections[setctionIndex];
            Table table = section.Tables[1] as Table;
            TableCell cell = table.Rows[rowIndex].Cells[colIndex];
            cell.Paragraphs[0].Text = value;
        }
        public void SetCellValue(int setctionIndex, int tableIndex, int rowIndex, int colIndex, string value)
        {
            rowIndex += 1;
            Section section = doc.Sections[setctionIndex];
            Table table = section.Tables[tableIndex] as Table;
            TableCell cell = table.Rows[rowIndex].Cells[colIndex];
            cell.Paragraphs[0].Text = value;
        }
        /// <summary>
        /// 获得表格的值
        /// </summary>
        /// <param name="rowIndex"></param>
        /// <param name="colIndex"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        public string GetCellValue(int rowIndex, int colIndex)
        {
            rowIndex += 6;
            Section section = doc.Sections[0];
            Table table = section.Tables[0] as Table;
            TableCell cell = table.Rows[rowIndex].Cells[colIndex];
            return cell.Paragraphs[0].Text;
        }
        /// <summary>
        /// 获得表格的值
        /// </summary>
        /// <param name="sectionIndex"></param>
        /// <param name="rowIndex"></param>
        /// <param name="colIndex"></param>
        /// <returns></returns>
        public string GetCellValue(int sectionIndex, int rowIndex, int colIndex)
        {
            rowIndex += 1;
            Section section = doc.Sections[sectionIndex];
            Table table = section.Tables[1] as Table;
            TableCell cell = table.Rows[rowIndex].Cells[colIndex];
            return cell.Paragraphs[0].Text;
        }
    }
}
