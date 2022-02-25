# Proxy Excel to JSON converter

A easy to use module which reads an excel file with given columns and converts it to json.

## Usage

```c#

 var excelConverter = new ExcelConverter();
 excelConverter.FilePath = "filename.xlsx";
 excelConverter.Columns = new List<string>()
            {
                "Id",
                "Name"
            };

 var json = await excelConverter.ToJson();
```

## Contributing
Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

Please make sure to update tests as appropriate.

## License
[MIT](https://choosealicense.com/licenses/mit/)