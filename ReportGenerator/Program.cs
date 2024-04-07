using DocumentFormat.OpenXml.Packaging;
using ReportComposer;
using System.ComponentModel.DataAnnotations;
using System.Text.Json;
using System.Text.Json.Nodes;

var templateFileName = @"TestWorkspace\resultTemplate.docx";
var dataFile = @"TestWorkspace\data.json";

var jsonData = System.IO.File.ReadAllText(dataFile);

var dstFile = @"TestWorkspace\result.docx";

var stream=new FileStream(dstFile, FileMode.Create);

var json = JsonSerializer.Deserialize<JsonNode>(jsonData);

if (json != null && json is JsonObject data)
{
	WordComposer composer = new WordComposer(templateFileName, data);

	composer.SaveToFile(stream);
}