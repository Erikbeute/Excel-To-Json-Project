using Newtonsoft.Json;
using OfficeOpenXml;
using VormerProject.Mappings;
using VormerProject.Models;
using VormerProject.Services;

public class ExcelService : IExcelService
{
    private readonly MappingConfig _mappingConfig;

    public ExcelService()
    {
        var configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "excelConfig.json");
        _mappingConfig = LoadMappingConfig(configPath);
    }

    public async Task<string> ReturnJson(IFormFile file)
    {
        var people = await ProcessExcelFile(file);
        var json = JsonConvert.SerializeObject(people);

        // gewoon 'hardcoded' pad voor de JSON - file. 
        var path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "../../../downloads"); 

        var filePath = Path.Combine(path, "excelToJsonOutput.json");
        await File.WriteAllTextAsync(filePath, json);
        return json;
    }



    private async Task<List<Person>> ProcessExcelFile(IFormFile file)
    {
        var people = new List<Person>();

        using (var stream = new MemoryStream())
        {
            await file.CopyToAsync(stream);
            using var package = new ExcelPackage(stream);
            var worksheet = package.Workbook.Worksheets[0];

            for (int row = 2; row <= worksheet.Dimension.Rows; row++) // start bij 2 because 1 is de header van de column
            {
                var person = new Person();

                foreach (var mapping in _mappingConfig.Mappings)
                {
                    var cellValue = worksheet.Cells[$"{mapping.ExcelColumn}{row}"].Text;
                    CellValueToPerson(mapping.ObjectField, cellValue, person);
                }

                people.Add(person);
            }
        }

        return people;
    }

    private static MappingConfig LoadMappingConfig(string filePath)
    {
        var json = File.ReadAllText(filePath);
        return JsonConvert.DeserializeObject<MappingConfig>(json);
    }

    private void CellValueToPerson(string objectField, string cellValue, Person person)
    {
        switch (objectField)
        {
            case "Voornaam":
                person.FirstName = cellValue;
                break;
            case "Achternaam":
                person.LastName = cellValue;
                break;
            case "Leeftijd":
                if (int.TryParse(cellValue, out int age))
                {
                    person.Age = age;
                }
                break;
            case "Functie":
                person.Function = cellValue;
                break;
        }
    }
}
