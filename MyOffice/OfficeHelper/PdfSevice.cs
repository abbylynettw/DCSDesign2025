
using Aspose.Pdf;
using Aspose.Pdf.Text;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WinformUI.OfficeHelper
{
    public class PdfService
    {
        private const string ASPOSE_LICENSE_TEXT =
            """
        <?xml version="1.0"?>
        <License>
        <Data>
        <LicensedTo>KMD A/S</LicensedTo>
        <EmailTo>iit-software@kmd.dk</EmailTo>
        <LicenseType>Site OEM</LicenseType>
        <LicenseNote>Up To 10 Developers And Unlimited Deployment Locations</LicenseNote>
        <OrderID>220815085749</OrderID>
        <UserID>324045</UserID>
        <OEM>This is a redistributable license</OEM>
        <Products>
        <Product>Aspose.Total for .NET</Product>
        </Products>
        <EditionType>Enterprise</EditionType>
        <SerialNumber>acea5afd-3c7d-4052-8991-f1e8522f63b4</SerialNumber>
        <SubscriptionExpiry>20230818</SubscriptionExpiry>
        <LicenseVersion>3.0</LicenseVersion>
        <LicenseInstructions>https://purchase.aspose.com/policies/use-license</LicenseInstructions>
        </Data>
        <Signature>d6CNxPzdmeo0I8EJmarUMRizSisbxluOwz5BdYKprWEyJbqjvs93//lCgP0tNzxIzvniD9T7PefYeEtlkQoVKV9fo3pdjfh2QrWFxJZuRby9yzfTqK7Ahghj81URDTpneve+RAL3Z63bwkCNH0anWR0Z1I6Bdug5L8QZpduoS5k=</Signature>
        </License>
        """;
        /// <summary>创建 <see cref="PdfService"/> 的新实例</summary>
        public PdfService()
        {
            new License().SetLicense(new MemoryStream(Encoding.UTF8.GetBytes(ASPOSE_LICENSE_TEXT)));
            Task.Run(() => new Document().Dispose());
        }

        /// <summary>
        /// 遍历所有页面替换文本
        /// </summary>
        /// <param name="inputPath">输入PDF文件路径</param>
        /// <param name="outputPath">输出PDF文件路径</param>
        /// <param name="oldText">要替换的文本</param>
        /// <param name="newText">替换后的文本</param>
        /// <returns>替换成功返回true，否则返回false</returns>
        public bool ReplaceTextOnAllPages(string inputPath, string outputPath, string oldText, string newText)
        {
            try
            {
                // 加载PDF文档
                Document pdfDocument = new Document(inputPath);

                // 获取总页数
                int pageCount = pdfDocument.Pages.Count;

                // 遍历每一页
                for (int pageNumber = 1; pageNumber <= pageCount; pageNumber++)
                {
                    // 创建TextFragmentAbsorber对象查找文本
                    TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber(oldText);

                    // 接受器选项
                    textFragmentAbsorber.TextSearchOptions = new TextSearchOptions(true);

                    // 在当前页面搜索文本
                    pdfDocument.Pages[pageNumber].Accept(textFragmentAbsorber);

                    // 获取找到的文本片段
                    TextFragmentCollection textFragmentCollection = textFragmentAbsorber.TextFragments;

                    // 遍历所有文本片段
                    foreach (TextFragment textFragment in textFragmentCollection)
                    {
                        // 更新文本
                        textFragment.Text = newText;
                    }
                }

                // 保存修改后的文档
                pdfDocument.Save(outputPath);

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"遍历替换页面文本时发生错误: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 遍历所有页面替换多个文本
        /// </summary>
        /// <param name="inputPath">输入PDF文件路径</param>
        /// <param name="outputPath">输出PDF文件路径</param>
        /// <param name="replacements">替换映射表，键为要替换的文本，值为替换后的文本</param>
        /// <returns>替换成功返回true，否则返回false</returns>
        public bool ReplaceMultipleTextsOnAllPages(string inputPath, string outputPath, Dictionary<string, string> replacements)
        {
            try
            {
                // 加载PDF文档
                Document pdfDocument = new Document(inputPath);

                // 获取总页数
                int pageCount = pdfDocument.Pages.Count;

                // 遍历每一页
                for (int pageNumber = 1; pageNumber <= pageCount; pageNumber++)
                {
                    // 遍历所有替换项
                    foreach (var replacement in replacements)
                    {
                        string oldText = replacement.Key;
                        string newText = replacement.Value;

                        // 创建TextFragmentAbsorber对象查找文本
                        TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber(oldText);

                        // 接受器选项
                        textFragmentAbsorber.TextSearchOptions = new TextSearchOptions(true);

                        // 在当前页面搜索文本
                        pdfDocument.Pages[pageNumber].Accept(textFragmentAbsorber);

                        // 获取找到的文本片段
                        TextFragmentCollection textFragmentCollection = textFragmentAbsorber.TextFragments;

                        // 遍历所有文本片段
                        foreach (TextFragment textFragment in textFragmentCollection)
                        {
                            // 更新文本
                            textFragment.Text = newText;
                        }
                    }
                }

                // 保存修改后的文档
                pdfDocument.Save(outputPath);

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"遍历替换多个文本时发生错误: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 遍历所有页面替换多个文本，支持自定义起始页码
        /// </summary>
        /// <param name="inputPath">输入PDF文件路径</param>
        /// <param name="outputPath">输出PDF文件路径</param>
        /// <param name="replacements">替换映射表，键为要替换的文本，值为替换后的文本</param>
        /// <param name="startPageNumber">起始页码，从这个数字开始编排页码</param>
        /// <returns>替换成功返回true，否则返回false</returns>
        public bool ReplaceMultipleTextsOnAllPages(string inputPath, string outputPath, Dictionary<string, string> replacements, Dictionary<int, string> bookmarkDic, int startPageNumber = 1)
        {
            try
            {
                // 加载PDF文档
                using Document pdfDocument = new Document(inputPath);

                // 获取总页数
                int pageCount = pdfDocument.Pages.Count;

                // 遍历每一页
                for (int pageNumber = 1; pageNumber <= pageCount; pageNumber++)
                {
                    // 创建页码替换字典
                    var pageReplacements = new Dictionary<string, string>(replacements);

                    // 计算当前页的显示页码
                    int displayPageNumber = startPageNumber + pageNumber - 1;

                    // 添加页码替换项
                    pageReplacements["YM1"] = $"{displayPageNumber}";
                    //pageReplacements["YM2"] = $"{pageCount}";
                    if (bookmarkDic[pageNumber].Contains("-")) bookmarkDic[pageNumber] = bookmarkDic[pageNumber].Split('-')[0];
                    pageReplacements["TUZHIBIANMA"] = $"{bookmarkDic[pageNumber]}";

                    // 遍历所有替换项
                    foreach (var replacement in pageReplacements)
                    {
                        string oldText = replacement.Key;
                        string newText = replacement.Value;

                        // 创建TextFragmentAbsorber对象查找文本
                        var textFragmentAbsorber = new TextFragmentAbsorber(oldText);

                        // 接受器选项
                        textFragmentAbsorber.TextSearchOptions = new TextSearchOptions(true);

                        // 在当前页面搜索文本
                        pdfDocument.Pages[pageNumber].Accept(textFragmentAbsorber);

                        // 获取找到的文本片段
                        TextFragmentCollection textFragmentCollection = textFragmentAbsorber.TextFragments;

                        // 遍历所有文本片段
                        foreach (TextFragment textFragment in textFragmentCollection)
                        {
                            // 更新文本
                            textFragment.Text = newText;
                        }
                    }
                }

                // 保存修改后的文档
                pdfDocument.Save(outputPath);
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"遍历替换多个文本时发生错误: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 修复版：替换特定页面的文本（您原代码中缺少参数和有语法错误）
        /// </summary>
        /// <param name="inputPath">输入PDF文件路径</param>
        /// <param name="outputPath">输出PDF文件路径</param>
        /// <param name="pageNumber">页码（从1开始）</param>
        /// <param name="oldText">要替换的文本</param>
        /// <param name="newText">替换后的文本</param>
        /// <returns>替换成功返回true，否则返回false</returns>
        public bool ReplaceTextOnPage(string inputPath, string outputPath, int pageNumber, string oldText, string newText)
        {
            try
            {
                // 加载PDF文档
                Document pdfDocument = new Document(inputPath);

                // 检查页码是否有效
                if (pageNumber < 1 || pageNumber > pdfDocument.Pages.Count)
                {
                    Console.WriteLine($"无效的页码: {pageNumber}");
                    return false;
                }

                // 创建TextFragmentAbsorber对象查找文本
                TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber(oldText);

                // 接受器选项
                textFragmentAbsorber.TextSearchOptions = new TextSearchOptions(true);

                // 只在指定页面搜索文本
                pdfDocument.Pages[pageNumber].Accept(textFragmentAbsorber);

                // 获取找到的文本片段
                TextFragmentCollection textFragmentCollection = textFragmentAbsorber.TextFragments;

                // 遍历所有文本片段
                foreach (TextFragment textFragment in textFragmentCollection)
                {
                    // 更新文本
                    textFragment.Text = newText;
                }

                // 保存修改后的文档
                pdfDocument.Save(outputPath);

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"替换特定页面文本时发生错误: {ex.Message}");
                return false;
            }
        }

        public Dictionary<int, string> GetBookmarks(string pdfPath,ref string cabinetName)
        {
            Dictionary<int, string> bookmarks = new Dictionary<int, string>();
            try
            {
                using (Document doc = new Document(pdfPath))
                {
                    int pageNumber = 1;  // 从第1页开始
                    foreach (OutlineItemCollection outlineItem in doc.Outlines)
                    {
                        if (string.IsNullOrEmpty(cabinetName)) cabinetName = outlineItem.Title;
                        foreach (OutlineItemCollection childItem in outlineItem)
                        {
                            // 按顺序添加书签标题
                            bookmarks.Add(pageNumber, childItem.Title);
                            pageNumber++;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"读取书签时出错: {ex.Message}");
            }
            return bookmarks;
        }

    }
}
