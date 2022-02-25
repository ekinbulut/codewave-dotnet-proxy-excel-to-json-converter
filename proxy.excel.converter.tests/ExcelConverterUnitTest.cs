using System;
using System.Collections.Generic;
using System.IO;
using Xunit;

namespace proxy.excel.converter.tests
{
    public class ExcelConverterUnitTest
    {
        [Fact]
        public async void Read_Value_From_Excel()
        {
            var excelConverter = new ExcelConverter();
            excelConverter.FilePath = "/Users/ekin/OneDrive/LibraryOS_Current.xlsx";
            excelConverter.Columns = new List<string>()
            {
                "Id",
                "Name",
                "Author",
                "Publisher",
                "Series",
                "Genre",
                "PublishDate",
                "No",
                "Cilt",
                "RackNumber",
                "Shelf",
                "CreationTime"
            };

            var expected = await excelConverter.ToJson();
            await File.WriteAllTextAsync("/Users/ekin/OneDrive/test.json", expected);

            Assert.NotNull(expected);
        }

        [Fact]
        public async System.Threading.Tasks.Task ToJson_FilePath_Is_Empty()
        {
            var excelConverter = new ExcelConverter
            {
                Columns = new List<string>()
                {
                    "Id",
                    "Name",
                    "Author",
                    "Publisher",
                    "Series",
                    "Genre",
                    "PublishDate",
                    "No",
                    "Cilt",
                    "RackNumber",
                    "Shelf",
                    "CreationTime"
                }
            };

            await Assert.ThrowsAsync<ArgumentNullException>(() => excelConverter.ToJson());

        }
        
        [Fact]
        public async System.Threading.Tasks.Task ToJson_Colum_Is_Less_Then_Default_Column()
        {
            var excelConverter = new ExcelConverter
            {
                FilePath =  "/Users/ekin/OneDrive/LibraryOS_Current.xlsx",
                Columns = new List<string>()
                {
                    "Id",
                    "Name",
                    "Author",
                    "Publisher",
                    "Series",
                    "Genre",
                    "PublishDate"
                }
            };

            await Assert.ThrowsAsync<Exception>(() => excelConverter.ToJson());

        }
    }
}