using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text.Json.Serialization;
using System.Threading.Tasks;
using ExcelDataReader;
using Newtonsoft.Json;

namespace proxy.excel.converter
{
    public class ExcelConverter
    {
        public string FilePath { get; set; }
        public List<string> Columns { get; set; }

        public ExcelConverter()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        }

        public Task<string> ToJson()
        {
            if (string.IsNullOrEmpty(FilePath))
            {
                throw new ArgumentNullException("FilePath can not be null");
            }

            using var fileStream = GetFileStream();
            var json = ReadFileStream(fileStream);
            return Task.FromResult(json);
        }

        private FileStream GetFileStream()
        {
            return File.Open(FilePath, FileMode.Open, FileAccess.Read);
        }

        private string ReadFileStream(FileStream fileStream)
        {
            using var reader = ExcelReaderFactory.CreateReader(fileStream);
            var dataset = reader.AsDataSet();
            ReplaceDefaultColumnNamesWithSelectedColumns(ref dataset);
            return JsonConvert.SerializeObject(dataset, Formatting.Indented);
        }

        private void ReplaceDefaultColumnNamesWithSelectedColumns(ref DataSet dataset)
        {
            foreach (DataTable dataTable in dataset.Tables)
            {
                if (dataTable.Columns.Count > Columns.Count)
                {
                    throw new Exception("Selected Columns must be more than file columns.");
                }
                        
                for (var i = 0; i < dataTable.Columns.Count; i++)
                {
                    dataTable.Columns[$"Column{i}"]!.ColumnName = Columns[i];
                }
            }
        }
    }
}