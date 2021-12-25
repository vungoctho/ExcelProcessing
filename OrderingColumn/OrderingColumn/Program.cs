// See https://aka.ms/new-console-template for more information
using OrderingColumn;

var folder = @"C:\Users\Tho.Vu\Desktop\KTTX3";
var files = Directory.GetFiles(folder);
foreach (var file in files)
{
    try
    {
        using (FileStream stream = File.Open(file, FileMode.Open, FileAccess.ReadWrite))
        {
            var excel = new OrderStudentNameProcess();
            var outFile = excel.ProcessFile(stream, folder, Path.GetFileName(file));
            Console.WriteLine($"Success to process {outFile}");
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error and SKIP file {file}");
    }
}

