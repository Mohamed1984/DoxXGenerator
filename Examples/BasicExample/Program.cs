using ReportComposer;
using System.Text.Json;
using System.Text.Json.Nodes;

var jsonData = System.IO.File.ReadAllText(@"data.json");

var json = JsonSerializer.Deserialize<JsonNode>(jsonData);

if (json != null && json is JsonObject data)
{
	WordComposer composer = new WordComposer(@"template.docx", data);

	composer.SaveToFile(@"result.docx");
}