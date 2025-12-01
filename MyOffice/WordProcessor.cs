using log4net;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace MyOffice
{
    /// <summary>
    /// Word文档处理器 - 高效实现版本(使用DocumentFormat.OpenXml)
    /// </summary>
    public class WordProcessor
    {
        private static readonly ILog Log = LogManager.GetLogger(typeof(WordProcessor));
        private StringBuilder _logBuilder;

        public WordProcessor(StringBuilder logBuilder = null)
        {
            this._logBuilder = logBuilder ?? new StringBuilder();
        }

        /// <summary>
        /// 处理Word文档 - 主方法
        /// </summary>
        public async Task<bool> ProcessWordFile(string filePath, List<ReplaceRule> replaceRules,
            List<OuterCodeAndTitle> outerCodeList, ProcessingOptions options)
        {
            if (!File.Exists(filePath))
            {
                LogMessage($"文件不存在: {filePath}", true);
                return false;
            }

            // 如果没有需要处理的选项，直接返回成功
            if (!NeedsProcessing(options))
                return true;

            LogMessage($"开始处理Word文档: {Path.GetFileName(filePath)}");       

            // 在后台线程处理文档
            bool success = await Task.Run(() => ProcessDocument(filePath, replaceRules, outerCodeList, options));

            LogMessage(success
                ? $"【成功】Word文档处理成功: {Path.GetFileName(filePath)}"
                : $"【失败】Word文档处理失败: {Path.GetFileName(filePath)}", !success);

            return success;
        }

        /// <summary>
        /// 检查是否需要进行处理
        /// </summary>
        private bool NeedsProcessing(ProcessingOptions options)
        {
            return options.项目名称修改 || options.项目编号修改 ||
                   options.内部编码修改 || options.外部编码修改 ||
                   options.页眉内部编码修改 || options.页眉版本状态修改 ||
                   options.页眉总页码修改 || options.页眉标题替换 ||
                   options.标题替换 || options.修改记录重置 ||
                   options.编校审批表格重置;
        }       
        /// <summary>
        /// 文档处理的核心方法
        /// </summary>
        private bool ProcessDocument(string filePath, List<ReplaceRule> replaceRules,
            List<OuterCodeAndTitle> outerCodeList, ProcessingOptions options)
        {
            try
            {
                using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true))
                {
                    LogMessage($"已打开文档: {Path.GetFileName(filePath)}");

                    // 获取文档标题并应用替换规则
                    string title = GetDocumentTitle(doc, replaceRules);

                    // 准备书签值字典
                    var bookmarkValues = PrepareBookmarkValues(title, outerCodeList, options, filePath);

                    // 更新文档中的所有书签
                    UpdateBookmarks(doc, bookmarkValues);

                    // 处理表格
                    ProcessTables(doc);

                    // 保存文档
                    doc.Save();
                }
                return true;
            }
            catch (Exception ex)
            {
                LogMessage($"处理文档出错: {Path.GetFileName(filePath)}, {ex.Message}\n{ex.StackTrace}", true);
                return false;
            }
        }

        /// <summary>
        /// 准备书签和值的字典
        /// </summary>
        private Dictionary<string, string> PrepareBookmarkValues(string title,
            List<OuterCodeAndTitle> outerCodeList, ProcessingOptions options, string filePath)
        {
            var bookmarkValues = new Dictionary<string, string>();

            // 添加各种书签的值
            if (options.项目名称修改) bookmarkValues["项目名称"] = options.项目名称值;
            if (options.项目编号修改) bookmarkValues["项目编号"] = options.项目编号值;
            if (options.内部编码修改)
            {
                bookmarkValues["内部编码"] = options.内部编码值;
                if (options.页眉内部编码修改)
                    bookmarkValues["页眉内部编码"] = options.内部编码值;
            }
            if (options.页眉总页码修改) bookmarkValues["页眉总页码"] = options.页眉总页码值;
            if (options.页眉版本状态修改) bookmarkValues["版本状态"] = options.页眉版本状态值;

            // 处理标题相关值
            if (options.标题替换 && !string.IsNullOrEmpty(title))
            {
                bookmarkValues["标题"] = title;
                if (options.页眉标题替换)
                    bookmarkValues["页眉标题"] = title;
            }

            // 处理外部编码
            if (options.外部编码修改 && !string.IsNullOrEmpty(title))
            {
                var matchedItem = outerCodeList?.FirstOrDefault(l => l.文件标题 == title);
                if (matchedItem != null && !string.IsNullOrEmpty(matchedItem.外部编码))
                {
                    bookmarkValues["外部编码"] = matchedItem.外部编码;
                }
                else
                {
                    LogMessage($"警告: 未找到标题'{title}'对应的有效外部编码: {Path.GetFileName(filePath)}");
                }
            }

            return bookmarkValues;
        }

        /// <summary>
        /// 获取文档标题
        /// </summary>
        private string GetDocumentTitle(WordprocessingDocument doc, List<ReplaceRule> replaceRules)
        {
            string title = string.Empty;

            try
            {
                // 首先尝试从内置文档属性中获取标题
                if (doc.CoreFilePropertiesPart != null)
                {
                    using (var stream = doc.CoreFilePropertiesPart.GetStream())
                    {
                        var xDoc = XDocument.Load(stream);
                        var dc = XNamespace.Get("http://purl.org/dc/elements/1.1/");
                        var titleElement = xDoc.Descendants(dc + "title").FirstOrDefault();
                        if (titleElement != null)
                            title = titleElement.Value;
                    }
                }

                // 如果文档属性中没有标题，尝试从书签中获取
                if (string.IsNullOrEmpty(title) && doc.MainDocumentPart != null)
                {
                    var bookmarkStart = doc.MainDocumentPart.Document.Descendants<BookmarkStart>()
                        .FirstOrDefault(b => b.Name == "标题");

                    if (bookmarkStart != null)
                    {
                        title = ExtractBookmarkText(bookmarkStart, doc.MainDocumentPart);
                    }
                }

                // 应用替换规则
                if (!string.IsNullOrEmpty(title) && replaceRules?.Count > 0)
                {
                    title = ApplyReplaceRules(title, replaceRules);
                }
            }
            catch (Exception ex)
            {
                LogMessage($"获取文档标题时出错: {ex.Message}", true);
            }

            return title;
        }

        /// <summary>
        /// 提取书签中的文本
        /// </summary>
        private string ExtractBookmarkText(BookmarkStart bookmarkStart, OpenXmlPart part)
        {
            try
            {
                // 查找对应的书签结束标记
                var bookmarkEnd = part.RootElement.Descendants<BookmarkEnd>()
                    .FirstOrDefault(b => b.Id == bookmarkStart.Id);

                if (bookmarkEnd == null) return string.Empty;

                // 收集书签开始和结束之间的所有文本
                var sb = new StringBuilder();
                OpenXmlElement current = bookmarkStart.NextSibling();

                while (current != null && current != bookmarkEnd)
                {
                    if (current is Run run)
                    {
                        foreach (var text in run.Descendants<Text>())
                        {
                            sb.Append(text.Text);
                        }
                    }
                    current = current.NextSibling();
                }

                return sb.ToString().Trim();
            }
            catch (Exception)
            {
                return string.Empty;
            }
        }

        /// <summary>
        /// 更新文档中的所有书签
        /// </summary>
        private void UpdateBookmarks(WordprocessingDocument doc, Dictionary<string, string> bookmarkValues)
        {
            if (bookmarkValues == null || bookmarkValues.Count == 0) return;

            try
            {
                // 处理主文档中的书签
                if (doc.MainDocumentPart != null)
                    UpdateBookmarksInPart(doc.MainDocumentPart, bookmarkValues, false);

                // 处理页眉中的书签
                if (doc.MainDocumentPart?.HeaderParts != null)
                {
                    foreach (var headerPart in doc.MainDocumentPart.HeaderParts)
                    {
                        UpdateBookmarksInPart(headerPart, bookmarkValues, true);
                    }
                }
            }
            catch (Exception ex)
            {
                LogMessage($"更新书签时出错: {ex.Message}", true);
            }
        }

        /// <summary>
        /// 在指定文档部分中更新书签
        /// </summary>
        private void UpdateBookmarksInPart(OpenXmlPart part, Dictionary<string, string> bookmarkValues, bool isHeader)
        {
            // 获取该部分中所有的书签开始标记
            var bookmarkStarts = part.RootElement.Descendants<BookmarkStart>()
                .Where(b => bookmarkValues.ContainsKey(b.Name)).ToList();

            foreach (var bookmarkStart in bookmarkStarts)
            {
                string bookmarkName = bookmarkStart.Name;
                string newValue = bookmarkValues[bookmarkName];

                // 查找对应的书签结束标记
                var bookmarkEnd = part.RootElement.Descendants<BookmarkEnd>()
                    .FirstOrDefault(b => b.Id == bookmarkStart.Id);

                if (bookmarkEnd == null) continue;

                // 提取当前书签文本
                string currentText = ExtractBookmarkText(bookmarkStart, part);

                // 如果内容没有变化，则跳过
                if (currentText == newValue) continue;

                // 更新书签内容
                ReplaceBookmarkContent(bookmarkStart, bookmarkEnd, newValue);

                string location = isHeader ? "页眉" : "文档";
                LogMessage($"已更新{location}书签 '{bookmarkName}': {currentText} -> {newValue}");
            }
        }

        /// <summary>
        /// 替换书签内容，同时保留原格式
        /// </summary>
        private void ReplaceBookmarkContent(BookmarkStart bookmarkStart, BookmarkEnd bookmarkEnd, string newText)
        {
            // 获取第一个Run元素以复制其格式
            Run templateRun = null;
            OpenXmlElement current = bookmarkStart.NextSibling();

            // 尝试找到一个Run元素作为格式模板
            while (current != null && current != bookmarkEnd && templateRun == null)
            {
                if (current is Run run)
                {
                    templateRun = run.CloneNode(true) as Run;
                    // 清除克隆的Run中的所有文本内容，但保留格式
                    templateRun.RemoveAllChildren<Text>();
                }
                current = current.NextSibling();
            }

            // 删除书签之间的所有内容
            current = bookmarkStart.NextSibling();
            while (current != null && current != bookmarkEnd)
            {
                var next = current.NextSibling();
                current.Remove();
                current = next;
            }

            // 创建新的文本节点，保留原有格式
            Run newRun;
            if (templateRun != null)
            {
                // 使用模板Run的格式
                newRun = templateRun;
                newRun.AppendChild(new Text(newText) { Space = SpaceProcessingModeValues.Preserve });
            }
            else
            {
                // 如果没有找到模板，创建一个新的Run
                newRun = new Run(new Text(newText) { Space = SpaceProcessingModeValues.Preserve });
            }

            // 插入新内容
            bookmarkStart.Parent.InsertAfter(newRun, bookmarkStart);
        }

        /// <summary>
        /// 处理表格
        /// </summary>
        private void ProcessTables(WordprocessingDocument doc)
        {
            try
            {
                if (doc.MainDocumentPart?.Document?.Body == null) return;

                // 获取文档中的所有表格
                var tables = doc.MainDocumentPart.Document.Body.Descendants<Table>().ToList();
                if (tables.Count == 0) return;

                // 处理第一个表格 - 修改第7行第1列
                if (tables.Count > 0)
                {
                    var firstTable = tables[0];
                    var rows = firstTable.Elements<TableRow>().ToList();

                    if (rows.Count >= 7) // 确保有足够的行
                    {
                        var row7 = rows[6]; // 第7行是索引6
                        var cells = row7.Elements<TableCell>().ToList();

                        if (cells.Count > 0) // 确保有至少一个单元格
                        {
                            string originalValue = string.Empty;

                            // 获取原始文本以便记录日志
                            if (cells[0].Elements<Paragraph>().FirstOrDefault() != null)
                            {
                                StringBuilder sb = new StringBuilder();
                                foreach (var text in cells[0].Descendants<Text>())
                                {
                                    sb.Append(text.Text);
                                }
                                originalValue = sb.ToString().Trim();
                            }

                            // 直接修改单元格内容为"A"
                            ModifyCellTextDirectly(cells[0], "A");

                            LogMessage($"已修改第一个表格第7行第1列单元格内容: '{originalValue}' -> 'A'");
                        }
                    }
                }

                // 处理第三个表格
                if (tables.Count >= 3)
                {
                    var thirdTable = tables[2];
                    var rows = thirdTable.Elements<TableRow>().ToList();

                    // 修改第2行第3列为空
                    if (rows.Count >= 2)
                    {
                        var row2 = rows[1]; // 第2行是索引1
                        var cells = row2.Elements<TableCell>().ToList();

                        if (cells.Count >= 3) // 确保有至少3个单元格
                        {
                            // 修改第3列单元格为空
                            ModifyCellTextDirectly(cells[2], string.Empty);
                        }
                    }

                    // 清空第3行及之后的所有单元格
                    for (int i = 2; i < rows.Count; i++)
                    {
                        foreach (var cell in rows[i].Elements<TableCell>())
                        {
                            // 直接清空单元格内容
                            ModifyCellTextDirectly(cell, string.Empty);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogMessage($"处理表格时出错: {ex.Message}", true);
            }
        }

        /// <summary>
        /// 直接修改单元格文本内容，保留原始格式结构
        /// </summary>
        private void ModifyCellTextDirectly(TableCell cell, string newContent)
        {
            try
            {
                // 查找单元格中的所有Text元素
                var textElements = cell.Descendants<Text>().ToList();

                if (textElements.Count > 0)
                {
                    // 保留第一个Text元素并修改其内容
                    var firstText = textElements[0];
                    firstText.Text = newContent;

                    // 清空其他Text元素（如果有的话）
                    for (int i = 1; i < textElements.Count; i++)
                    {
                        textElements[i].Text = string.Empty;
                    }
                }
                else
                {
                    // 如果单元格中没有Text元素，找出第一个段落或创建一个
                    var paragraph = cell.Elements<Paragraph>().FirstOrDefault();
                    if (paragraph == null)
                    {
                        paragraph = new Paragraph();
                        cell.AppendChild(paragraph);
                    }

                    // 查找段落中的第一个Run或创建一个
                    var run = paragraph.Elements<Run>().FirstOrDefault();
                    if (run == null)
                    {
                        run = new Run();
                        paragraph.AppendChild(run);
                    }

                    // 创建一个新的Text元素
                    var text = new Text(newContent) { Space = SpaceProcessingModeValues.Preserve };

                    // 检查Run中是否已有Text元素
                    var existingText = run.Elements<Text>().FirstOrDefault();
                    if (existingText != null)
                    {
                        // 替换现有Text元素
                        existingText.Text = newContent;
                    }
                    else
                    {
                        // 添加新的Text元素
                        run.AppendChild(text);
                    }
                }
            }
            catch (Exception ex)
            {
                LogMessage($"直接修改单元格文本时出错: {ex.Message}", true);
            }
        }

        /// <summary>
        /// 应用替换规则到文本
        /// </summary>
        private string ApplyReplaceRules(string text, List<ReplaceRule> rules)
        {
            if (string.IsNullOrEmpty(text) || rules == null || rules.Count == 0)
                return text;

            string result = text;

            foreach (var rule in rules)
            {
                if (string.IsNullOrEmpty(rule.FindText)) continue;

                try
                {
                    switch (rule.Type)
                    {
                        case RuleType.包含:
                            result = result.Replace(rule.FindText, rule.ReplaceText);
                            break;
                        case RuleType.相等:
                            if (result == rule.FindText)
                                result = rule.ReplaceText;
                            break;
                    }
                }
                catch (Exception ex)
                {
                    LogMessage($"应用替换规则时出错: {rule.FindText} -> {rule.ReplaceText}, {ex.Message}", true);
                }
            }

            return result;
        }

        /// <summary>
        /// 记录日志消息
        /// </summary>
        private void LogMessage(string message, bool isError = false)
        {
            if (_logBuilder != null)
                _logBuilder.AppendLine(message);

            if (isError)
                Log.Error(message);
            else
                Log.Info(message);
        }
    }
}