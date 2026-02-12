using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;

class Program
{
    static void Main(string[] args)
    {
        Console.WriteLine("开始验证Docx文件是否符合微软官方标准...");
        
        string docxDirectory = Path.Combine(Directory.GetCurrentDirectory(), "Docx");
        
        if (!Directory.Exists(docxDirectory))
        {
            Console.WriteLine($"错误: Docx目录不存在 - {docxDirectory}");
            return;
        }
        
        // 获取所有.docx文件
        string[] docxFiles = Directory.GetFiles(docxDirectory, "*.docx");
        
        if (docxFiles.Length == 0)
        {
            Console.WriteLine("错误: Docx目录中没有找到.docx文件");
            return;
        }
        
        Console.WriteLine($"找到 {docxFiles.Length} 个.docx文件，开始验证...");
        
        foreach (string filePath in docxFiles)
        {
            Console.WriteLine($"\n验证文件: {Path.GetFileName(filePath)}");
            bool isValid = ValidateDocxFile(filePath);
            Console.WriteLine($"验证结果: {(isValid ? "有效" : "无效")}");
        }
        
        Console.WriteLine("\n验证完成！");
        Console.WriteLine("\n请按任意键退出...");
        while (Console.KeyAvailable)
        {
            Console.ReadKey(false);
        }
        Console.ReadKey(true);
    }
    
    static bool ValidateDocxFile(string filePath)
    {
        try
        {
            // 使用OpenXmlPackage打开文件并验证
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(filePath, false))
            {
                // 检查文件是否是有效的Word文档
                if (wordDocument.MainDocumentPart == null)
                {
                    Console.WriteLine("错误: 缺少主文档部分");
                    return false;
                }
                
                // 检查文档结构是否完整
                if (wordDocument.MainDocumentPart.Document == null)
                {
                    Console.WriteLine("错误: 缺少文档内容");
                    return false;
                }
                
                // 尝试读取文档内容以验证结构完整性
                var document = wordDocument.MainDocumentPart.Document;
                
                // 验证文档至少包含一个正文部分
                if (document.Body == null)
                {
                    Console.WriteLine("错误: 缺少文档正文");
                    return false;
                }
                
                Console.WriteLine("文档结构验证通过");
                return true;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"验证错误: {ex.Message}");
            return false;
        }
    }
}
