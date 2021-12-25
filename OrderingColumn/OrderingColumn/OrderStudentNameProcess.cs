using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace OrderingColumn
{
    internal class Student
    {
        public string Ho { get; set; }
        public string Ten { get; set; }
        public string Lot { get; set; }
        public int Mark { get; set; }
    }
    internal class OrderStudentNameProcess
    {
        private readonly int _nameColumn = 2; // B
        public string ProcessFile(Stream stream, string folder, string filename)
        {
            using (var wbook = new XLWorkbook(stream))
            {
                var sheet = wbook.Worksheet(1);
                var rowIndex = 1;
                int _markColumn = -1;
                List<Student> students = new List<Student>();
                foreach (var row in sheet.Rows())
                {
                    try
                    {
                        if (rowIndex == 1)
                        {
                            // header row => read the column names
                            var col = 1;
                            foreach (var cell in row.Cells())
                            {
                                if (string.Equals("Điểm", cell.Value.ToString(),
                                    StringComparison.InvariantCultureIgnoreCase))
                                {
                                    _markColumn = col;
                                    break;
                                }
                                col++;
                            }
                            if (_markColumn <= 0) throw new Exception("NO FOUND OUT Điểm COLUMN");
                        }
                        else
                        {
                            var student = MappingRowToStudent(row, _markColumn);
                            if (student != null)
                                students.Add(student);
                            else break;
                        }
                        rowIndex++;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"ERROR at row {rowIndex} in {folder}/{filename}: {ex.Message}");
                        throw;
                    }
                }
                CultureInfo cul = CultureInfo.GetCultureInfo("vi-VN");
                students = students.OrderBy(s => s.Ten, StringComparer.Create(cul, false)).ThenBy(s => s.Ho, StringComparer.Create(cul, false)).ThenBy(s => s.Lot, StringComparer.Create(cul, false)).ToList();
                var worksheet = wbook.Worksheets.Add("Ordered");
                int i = 1;
                foreach (var student in students)
                {
                    worksheet.Cells($"A{i}").Value = $"{student.Ho} {student.Lot} {student.Ten}";
                    worksheet.Cells($"B{i}").Value = student.Mark;
                    i++;
                }

                var outFile = $"{folder}/Ordered_{filename}";
                wbook.SaveAs(outFile);
                return outFile;
            }
        }

        public Student MappingRowToStudent(IXLRow row, int markColumn)
        {
            var name = row.Cell(_nameColumn).Value.ToString();
            var mark = int.Parse(row.Cell(markColumn).Value.ToString());
            if(!string.IsNullOrEmpty(name))
            {
                return CreateStudent(name, mark);
            }
            return null;
        }

        public  Student CreateStudent(string? name, int mark)
        {
            var items = name.Split(" ");
            return new Student()
            {
                Ho = items[0],
                Ten = items[items.Length - 1],
                Lot = string.Join(' ', items.Skip(1).Take(items.Length - 2).ToArray()),
                Mark = mark,
            };
        }
    }
}
