using log4net;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyOffice.OfficeHelper
{
    public  class LicenseManager
    {   // 获取日志记录器     
        private static readonly ILog Log = MyOffice.LogHelper.LogManager.GetLogger<LicenseManager>();
        private const string ASPOSE_LICENSE_TEXT = @"<?xml version=""1.0""?>
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
        </License>";
    
        public  void SetLicense()
        {         
            try
            {
               
                // 输出详细的环境信息，有助于排查问题               
                Log.Info ("=== 环境信息 ===");
                Log.Info($"操作系统: {Environment.OSVersion}");
                Log.Info($".NET 版本: {Environment.Version}");
                Log.Info($"64位进程: {Environment.Is64BitProcess}");
                Log.Info($"64位操作系统: {Environment.Is64BitOperatingSystem}");
                Log.Info($"当前目录: {Environment.CurrentDirectory}");
                Log.Info($"机器名: {Environment.MachineName}");
                Log.Info($"用户名: {Environment.UserName}");
                Log.Info($"系统目录: {Environment.SystemDirectory}");
                Log.Info($"当前时间: {DateTime.Now}");

                // 记录Aspose版本信息
                Log.Info($"Aspose.Words 版本: {typeof(Aspose.Words.Document).Assembly.GetName().Version}");
              
              



                // 记录正在尝试加载许可证
                Log.Info("正在尝试加载 Aspose 许可证...");
               

                // 将许可证字符串转换为内存流
                using (MemoryStream licenseStream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(ASPOSE_LICENSE_TEXT)))
                {
                    // 记录许可证内容的哈希值（用于确认内容一致性）
                    byte[] hash = System.Security.Cryptography.SHA256.Create().ComputeHash(licenseStream);
                    string hashString = BitConverter.ToString(hash).Replace("-", "");
                    Log.Info($"许可证内容哈希: {hashString}");
                 

                    // 重置流位置
                    licenseStream.Position = 0;

                    // 尝试加载许可证
                    var license = new Aspose.Words.License();
                    license.SetLicense(licenseStream);

                    // 添加日志记录成功加载许可证
                    Log.Info("成功从内存加载 Aspose 许可证");
                  

                    // 尝试创建一个简单的文档来验证许可证是否真正有效
                    try
                    {
                        var doc = new Aspose.Words.Document();
                        doc.FirstSection.Body.FirstParagraph.AppendChild(new Aspose.Words.Run(doc, "测试文档"));
                        using (MemoryStream testStream = new MemoryStream())
                        {
                            doc.Save(testStream, Aspose.Words.SaveFormat.Docx);
                            Log.Info("成功创建测试文档，许可证有效");
                           
                        }
                    }
                    catch (Exception testEx)
                    {
                        Log.Info($"许可证虽然加载成功，但无法正常使用: {testEx.Message}");
                        Log.Info($"内部异常: {testEx.InnerException?.Message ?? "无"}");
                       
                    }
                }
            }
            catch (Exception ex)
            {
                // 记录详细的异常信息
                StringBuilder errorDetails = new StringBuilder();
                errorDetails.AppendLine("=== Aspose 许可证加载失败 ===");
                errorDetails.AppendLine($"异常类型: {ex.GetType().FullName}");
                errorDetails.AppendLine($"异常消息: {ex.Message}");

                // 记录内部异常链（通常包含更详细的原因）
                Exception innerEx = ex.InnerException;
                int level = 1;
                while (innerEx != null)
                {
                    errorDetails.AppendLine($"内部异常 {level}: {innerEx.GetType().FullName}");
                    errorDetails.AppendLine($"内部异常 {level} 消息: {innerEx.Message}");
                    errorDetails.AppendLine($"内部异常 {level} 堆栈: {innerEx.StackTrace}");
                    innerEx = innerEx.InnerException;
                    level++;
                }

                errorDetails.AppendLine($"异常堆栈: {ex.StackTrace}");

                // 尝试获取更多Aspose相关的诊断信息
                try
                {
                    var asposeAssembly = typeof(Aspose.Words.Document).Assembly;
                    errorDetails.AppendLine($"Aspose.Words 程序集位置: {asposeAssembly.Location}");
                    errorDetails.AppendLine($"Aspose.Words 程序集版本: {asposeAssembly.GetName().Version}");

                    // 检查是否有冲突的程序集
                    var loadedAssemblies = AppDomain.CurrentDomain.GetAssemblies()
                        .Where(a => a.GetName().Name.StartsWith("Aspose"))
                        .Select(a => $"{a.GetName().Name} v{a.GetName().Version} 位置: {a.Location}")
                        .ToList();

                    if (loadedAssemblies.Count > 0)
                    {
                        errorDetails.AppendLine("已加载的Aspose程序集:");
                        foreach (var assembly in loadedAssemblies)
                        {
                            errorDetails.AppendLine($"  - {assembly}");
                        }
                    }
                }
                catch (Exception diagnosticEx)
                {
                    errorDetails.AppendLine($"获取诊断信息时出错: {diagnosticEx.Message}");
                }

                // 将详细错误信息写入日志
                string errorInfo = errorDetails.ToString();
                Log.Info(errorInfo);
              

                // 同时写入文件以便后续分析
                try
                {
                    string logPath = Path.Combine(Environment.CurrentDirectory, "Aspose_License_Error.log");
                    File.WriteAllText(logPath, errorInfo);
                    Log.Info($"详细错误日志已写入: {logPath}");
                }
                catch
                {
                    // 忽略写入文件的错误
                }
            }
        }
    }
}
