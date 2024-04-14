using System.Net.Security;
using System.Text;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text.RegularExpressions;
using System.Text.Json.Nodes;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using System.Diagnostics;
using System.Collections.Generic;
using System.Linq;
using System.IO;

namespace ReportComposer
{
	internal enum HandleState
	{
		Advance,SeekToMatchingEndRepeat,Error
	}
	public class WordComposer
	{
		public WordComposer(string templateFilePath, JsonObject jsonData)
		{
			_template = WordprocessingDocument.Open(templateFilePath, false)
				.Clone(memStream, true);

			contextStack.Push(new ContextEntry(jsonData, "$"));
		}

		private MemoryStream memStream = new MemoryStream();



		public void SaveToFile(string fileName)
		{
			var stream = new FileStream(fileName, FileMode.Create);
			SaveToFile(stream);
			stream.Close();
		}

		public void SaveToFile(Stream file)
		{
			injectData();
			_template.Save();
			memStream.Position = 0;
			memStream.CopyTo(file);
		}


		private void injectData()
		{
			Debug.Assert(_template?.MainDocumentPart != null);

			var doc = _template.MainDocumentPart.Document;

			Debug.Assert(doc.FirstChild != null);
			Debug.Assert(doc.LastChild != null);

			processNodeRange(doc.FirstChild, doc.LastChild);

			foreach (HeaderPart header in _template.MainDocumentPart.HeaderParts)
			{
				Debug.Assert(header.Header.FirstChild != null);
				Debug.Assert(header.Header.LastChild != null);
				processNodeRange(header.Header.FirstChild, header.Header.LastChild);
			}
			foreach(FooterPart footer in _template.MainDocumentPart.FooterParts)
			{
				Debug.Assert(footer.Footer.FirstChild != null);
				Debug.Assert(footer.Footer.LastChild != null);
				processNodeRange(footer.Footer.FirstChild, footer.Footer.LastChild);
			}
		}

		private void processNodeRange(OpenXmlElement firstNode, OpenXmlElement lastNode)
		{
			HandleState state = HandleState.Advance;

			Queue<SdtElement> directives = new Queue<SdtElement>();

			var node = firstNode;
			while (true)
			{
				if (node is SdtElement sdt)
					directives.Enqueue(sdt);
				else
					foreach (var e in node.Descendants<SdtElement>())
					{
						directives.Enqueue(e);
					}

				if (node == lastNode)
					break;
				node = node.NextSibling();
				if (node == null)
					break;
			}

			foreach (var directive in directives)
			{
				//check if the directive element is deleted due to previous operations
				if (isDetached(directive))
					continue;
				state = handleDirective(directive, state);

			}
		}

		private bool isDetached(OpenXmlElement? e)
		{
			OpenXmlElement? p=null;
			while(e!=null)
			{
				p = e;
				e = e.Parent;
			}

			return p != _template.MainDocumentPart?.Document
				&& (!_template?.MainDocumentPart?.HeaderParts.Any(e => e.Header == p)??true)
				&& (!_template?.MainDocumentPart?.FooterParts.Any(e => e.Footer == p)??true);

		}

		

		private int repeatDepth = 0;
		private HandleState handleDirective(SdtElement cell, HandleState state)
		{
			//if in error skip all rest directives
			if (state == HandleState.Error)
				return state;

			Debug.Assert(cell.Parent != null);
			var directive = cell.InnerText.Trim();

			if (state == HandleState.SeekToMatchingEndRepeat)
			{
				if (directive.StartsWith("@EndRepeat") && repeatDepth == 0)
					return handleEndRepeat(cell);
				else if (directive.StartsWith("@Repeat"))
					repeatDepth++;
				else if (directive.StartsWith("@EndRepeat") && repeatDepth > 0)
					repeatDepth--;
				if (repeatDepth < 0)
					return HandleState.Error;
				return state;
			}

			if (directive.StartsWith("@") == false)
				return HandleState.Advance;

			var regBind = new Regex(@"@(\[(?<contextName>\w(\w|\d)*)\])?{(?<bindExpr>.+)}");
			var match = regBind.Match(directive);
			if (match.Success)
			{
				var context = getContext();
				var nameGroup = match.Groups["contextName"];
				if (nameGroup.Success)
				{
					var name = nameGroup.Value;
					context = getContext(name);
				}
				var bindExpr = match.Groups["bindExpr"].Value;
				if (context != null)
					return handleBindExpr(cell, bindExpr, context);
				else
					return HandleState.Advance;
			}

			var regContext = new Regex(@"@Context(\[(?<contextName>\w(\w|\d)*)\])?{(?<bindExpr>.+)}");

			match = regContext.Match(directive);
			if (match.Success)
			{
				string? newContextName = null;
				var newContextNameGroup = match.Groups["contextName"];
				if (newContextNameGroup.Success)
					newContextName = newContextNameGroup.Value;

				var newContextPath = match.Groups["bindExpr"]?.Value ?? "";

				return handleStartContext(cell, newContextName, newContextPath);
			}

			var regEndContext = new Regex(@"@EndContext");
			match = regEndContext.Match(directive);
			if (match.Success && contextStack.Count > 0)
			{
				return handleEndContext(cell);
			}

			var regRepeat = new Regex(@"@Repeat{(?<bindExpr>.+)}");
			match = regRepeat.Match(directive);
			if (match.Success)
			{
				string bindExpr = match.Groups["bindExpr"].Value;
				var context = getContext();
				if (context != null)
					return handleRepeat(cell, context, bindExpr);
				else
					return HandleState.Advance;
			}

			var regRowRepeat = new Regex(@"@RowRepeat\[(?<indexVar>\w(\w|\d)*)(\,(?<rowSpan>\d+))?\]{(?<bindArray>.+)}");
			match = regRowRepeat.Match(directive);
			if (match.Success)
			{
				var indexVar = match.Groups["indexVar"]?.Value ?? "";
				var rowSpanStr = match.Groups["rowSpan"].Success ?
					match.Groups["rowSpan"].Value : "1";
				var rowSpan = int.Parse(rowSpanStr);
				var bindArrayName = match.Groups["bindArray"]?.Value;
				if (contextStack.Count > 0 && bindArrayName != null)
				{
					var context = contextStack.Peek();
					var bindArr = GetBindingExprValue(context.Context, bindArrayName);

					if (bindArr is JsonArray data)
					{
						return handleRowRepeat(cell, indexVar, data.Count, rowSpan);
					}
				}
				return HandleState.Advance;
			}

			var regColRepeat = new Regex(@"@ColRepeat\[(?<indexVar>\w(\w|\d)*)(\,(?<colSpan>\d+))?\]{(?<bindArray>.+)}");
			match = regColRepeat.Match(directive);
			if (match.Success)
			{
				var indexVar = match.Groups["indexVar"]?.Value ?? "";
				var colSpanStr = match.Groups["colSpan"].Success ?
					match.Groups["colSpan"].Value : "1";
				var colSpan = int.Parse(colSpanStr);
				var bindArrayName = match.Groups["bindArray"]?.Value;
				if (contextStack.Count > 0 && bindArrayName != null)
				{
					var context = contextStack.Peek();
					var bindArr = GetBindingExprValue(context.Context, bindArrayName);

					if (bindArr is JsonArray data)
					{
						return handleColRepeat(cell, indexVar, data.Count, colSpan);
					}
				}
				return HandleState.Advance;
			}

			var regShowRow = new Regex(@"@RowShow{(?<maskVar>.+)}");
			match = regShowRow.Match(directive);
			if (match.Success)
			{
				var maskVar = match.Groups["maskVar"]?.Value ?? "";
				
				if (contextStack.Count > 0)
				{
					var context = contextStack.Peek();
					var flag = GetBindingExprValue(context.Context, maskVar);

					if (flag is JsonValue val)
					{
						if ((bool?)flag == false)
						{
							return handleRemoveRow(cell);
						}
						else
							RemovePlaceholder(cell);
					}
				}
				return HandleState.Advance;
			}

			var regHideRow = new Regex(@"@RowHide{(?<maskVar>.+)}");
			match = regHideRow.Match(directive);
			if (match.Success)
			{
				var maskVar = match.Groups["maskVar"]?.Value ?? "";

				if (contextStack.Count > 0)
				{
					var context = contextStack.Peek();
					var flag = GetBindingExprValue(context.Context, maskVar);

					if (flag is JsonValue val)
					{
						if ((bool?)flag == true)
						{
							return handleRemoveRow(cell);
						}
						else
							RemovePlaceholder(cell);
					}
				}
				return HandleState.Advance;
			}


			var regShowCol = new Regex(@"@ColShow{(?<maskVar>.+)}");
			match = regShowCol.Match(directive);
			if (match.Success)
			{
				var maskVar = match.Groups["maskVar"]?.Value ?? "";

				if (contextStack.Count > 0)
				{
					var context = contextStack.Peek();
					var flag = GetBindingExprValue(context.Context, maskVar);
					if (flag is JsonValue val)
					{
						if ((bool?)flag == false)
						{
							return handleRemoveCol(cell);
						}
						else
							RemovePlaceholder(cell);
					}
				}
				return HandleState.Advance;
			}

			var regHideCol = new Regex(@"@ColHide{(?<maskVar>.+)}");
			match = regHideCol.Match(directive);
			if (match.Success)
			{
				var maskVar = match.Groups["maskVar"]?.Value ?? "";

				if (contextStack.Count > 0)
				{
					var context = contextStack.Peek();
					var flag = GetBindingExprValue(context.Context, maskVar);

					if (flag is JsonValue val)
					{
						if ((bool?)flag == true)
						{
							return handleRemoveCol(cell);
						}
						else
							RemovePlaceholder(cell);
					}
				}
				return HandleState.Advance;
			}

			return HandleState.Advance;
		}

		private HandleState handleRemoveRow(OpenXmlElement placeholder)
		{
			OpenXmlElement? row = placeholder;
			while(row is not TableRow)
			{
				if (row?.Parent == null)
					break;
				row = row.Parent;
			}

			row?.Remove();
			return HandleState.Advance;
		}

		private HandleState handleRemoveCol(OpenXmlElement placeholder)
		{
			OpenXmlElement? e = placeholder;
			OpenXmlElement? rowElem = null;
			OpenXmlElement? cellElem = null;
			while (!(e is TableRow) && e != null)
			{
				cellElem = e;
				rowElem = e.Parent;
				e = e.Parent;
			}
			var row = rowElem as TableRow;

			if (row == null)
				return HandleState.Advance;

			var firstRowOffset = 0;
			for (firstRowOffset = 0; firstRowOffset < row.ChildElements.Count; firstRowOffset++)
				if (row.ChildElements[firstRowOffset] == cellElem)
					break;
			if (firstRowOffset == row.ChildElements.Count)
				return HandleState.Advance;

			//column index may not be equal to firstRowOffset because noncell
			//elements exist
			int colIndex = 0;
			for (int k = 0; k < firstRowOffset; k++)
			{
				var elm = row.ChildElements[k];
				if (elm is TableCell || elm is SdtCell)
					colIndex++;
			}

			var rowContainer = row.Parent;

			Debug.Assert(rowContainer != null);

			foreach (var item in rowContainer)
			{
				if (item is TableRow)
				{
					var cellCounter = 0;
					foreach (var x in item.ChildElements)
					{
						if (x is TableCell || x is SdtCell)
						{
							if (cellCounter == colIndex)
							{
								x.Remove();
								break;
							}
							cellCounter++;
						}
					}
				}
			}
			return HandleState.Advance;
		}

		private void RemovePlaceholder(OpenXmlElement placeholder)
		{
			if (placeholder is SdtCell)
			{//if it is a cell, we can't simply remove it !
				placeholder.InsertAfterSelf(new TableCell(
					new Paragraph(new Run(new Text(" ")))));
			}
			placeholder.Remove();
		}
		private HandleState handleColRepeat(SdtElement placeholder, string indexLabel, int dataLen, int colSpan = 1)
		{
			//get the parent row
			OpenXmlElement? e = placeholder;
			OpenXmlElement? rowElem = null;
			OpenXmlElement? cellElem = null;
			while (!(e is TableRow) && e != null)
			{
				cellElem = e;
				rowElem = e.Parent;
				e = e.Parent;
			}

			var row = rowElem as TableRow;

			if (row == null)
				return HandleState.Advance;

			var firstRowOffset = 0;
			for (firstRowOffset = 0; firstRowOffset < row.ChildElements.Count; firstRowOffset++)
				if (row.ChildElements[firstRowOffset] == cellElem)
					break;
			if (firstRowOffset == row.ChildElements.Count || firstRowOffset + colSpan > row.ChildElements.Count)
				return HandleState.Advance;

			//column index may not be equal to firstRowOffset because noncell
			//elements exist
			int colIndex = 0;
			for (int k = 0; k < firstRowOffset; k++)
			{
				var elm = row.ChildElements[k];
				if (elm is TableCell || elm is SdtCell)
					colIndex++;
			}
			//now remove the placeholder
			RemovePlaceholder(placeholder);

			var rowContainer = row.Parent;

			Debug.Assert(rowContainer != null);

			foreach(var item in rowContainer)
			{
				OpenXmlElement? firstInsertedCell = null;
				OpenXmlElement? lastInsertedCell = null;
				if (item is TableRow r)
				{
					//we search for the first cell by
					//counting colIndex cells from row start
					OpenXmlElement? firstTemplateCell = null;
					var cellCounter = 0;
					foreach (var x in r.ChildElements)
					{
						if (x is TableCell || x is SdtCell)
						{
							if (cellCounter == colIndex)
							{
								firstTemplateCell = x;
								break;
							}
							cellCounter++;
						}
					}

					if (firstTemplateCell == null)//if not enough cells, ignore this row
						continue;

					for (int j = 0; j < dataLen; j++)
					{
						var templateCell = firstTemplateCell;
						Debug.Assert(templateCell != null);
						for (int i = 0; i < colSpan; i++)
						{
							Debug.Assert(templateCell != null);
							var newCell = templateCell.CloneNode(true);
							firstTemplateCell.InsertBeforeSelf(newCell);
							InjectIndex(newCell, indexLabel, j);
							templateCell = templateCell.NextSibling();

							if (i == 0 && j == 0)
								firstInsertedCell = newCell;

							lastInsertedCell = newCell;
						}
					}
					//remove template cells
					var rtemplateCell = firstTemplateCell;
					for (int i = 0; i < colSpan; i++)
					{
						Debug.Assert(rtemplateCell != null);

						var oldTemplateCell = rtemplateCell;
						rtemplateCell = rtemplateCell.NextSibling();

						oldTemplateCell.Remove();
					}
					if (firstInsertedCell != null)
					{
						Debug.Assert(lastInsertedCell != null);
						processNodeRange(firstInsertedCell, lastInsertedCell);
					}
				}
			}

			//placeholder.Remove();
			//if (e is TableRow row)
			//{
			//	List<TableRow> dataRows = new List<TableRow>();
			//	dataRows.Add(row);

			//	int rowCounter = 0;
			//	while (row.NextSibling<TableRow>() is TableRow nextRow && rowCounter < rowSpan)
			//	{
			//		dataRows.Add(nextRow);
			//		row = nextRow;
			//		rowCounter++;
			//	}

			//	TableRow? firstNewRow = null;
			//	TableRow? lastNewRow = null;
			//	//now we have all rows to work on
			//	for (int i = 0; i < dataLen; i++)
			//	{

			//		for (int j = 0; j < dataRows.Count; j++)
			//		{
			//			var r = dataRows[j].CloneNode(true) as TableRow;

			//			Debug.Assert(r != null);

			//			dataRows[0].InsertBeforeSelf(r);

			//			InjectRowIndex(r, indexLabel, i);

			//			if (i == 0 && j == 0)
			//				firstNewRow = r;
			//			lastNewRow = r;
			//		}
			//	}
			//	//delete original template data rows
			//	foreach (var r in dataRows)
			//	{
			//		r.Remove();
			//	}
			//	if (firstNewRow != null)
			//	{
			//		Debug.Assert(lastNewRow != null);
			//		processNodeRange(firstNewRow, lastNewRow);
			//	}
			//}
			return HandleState.Advance;
		}


		private HandleState handleRowRepeat(SdtElement placeholder,string indexLabel,int dataLen,int rowSpan=1)
		{
			//get the Row
			OpenXmlElement? e = placeholder;
			while(!(e is TableRow)&& e!=null)
			{
				e = e.Parent;
			}
			placeholder.Remove();
			if(e is TableRow row)
			{
				List<TableRow> dataRows = new List<TableRow>();
				dataRows.Add(row);

				int rowCounter = 0;
				while(row.NextSibling<TableRow>()is TableRow nextRow && rowCounter<rowSpan)
				{
					dataRows.Add(nextRow);
					row = nextRow;
					rowCounter++;
				}

				TableRow? firstNewRow=null;
				TableRow? lastNewRow=null;
				//now we have all rows to work on
				for(int i=0;i<dataLen;i++)
				{
					
					for(int j=0;j<dataRows.Count;j++)
					{
						var r = dataRows[j].CloneNode(true) as TableRow;

						Debug.Assert(r != null);

						dataRows[0].InsertBeforeSelf(r);

						InjectIndex(r, indexLabel, i);

						if (i == 0 && j == 0)
							firstNewRow = r;
						lastNewRow = r;
					}
				}
				//delete original template data rows
				foreach (var r in dataRows)
				{
					r.Remove();
				}
				if (firstNewRow != null)
				{
					Debug.Assert(lastNewRow != null);
					processNodeRange(firstNewRow, lastNewRow);
				}
			}
			return HandleState.Advance;
		}

		private void InjectIndex(OpenXmlElement elem, string indexlabel, int indexValue)
		{
			foreach (var sdt in elem.Descendants<SdtElement>())
			{
				var str = sdt.InnerText;
				str = str.Replace($"#{indexlabel}", "" + indexValue);
				injectContentInSdt(sdt, str);
			}
		}
		private RepeatData? _curRepeatData=null;
		private HandleState handleRepeat(SdtElement startPlaceholder,JsonObject context, string bindExpr)
		{
			var arr=GetBindingExprValue(context, bindExpr);
			if (arr is JsonArray data)
			{
				_curRepeatData = new RepeatData(startPlaceholder,bindExpr, data);
				return HandleState.SeekToMatchingEndRepeat;
			}
			return HandleState.Advance;
		}
		private HandleState handleEndRepeat(SdtElement endPlaceholder)
		{
			Debug.Assert(_curRepeatData != null);

			var startNode = _curRepeatData.StartPlaceholder;
			var endNode = endPlaceholder;

			var data = _curRepeatData.Data;

			var innerContent = getRepeatInnerElements(startNode, endNode,out var commonParent);

			Debug.Assert(commonParent != null);

			var firstElement = innerContent[0];
			var lastElement = innerContent[innerContent.Count - 1];
			//make the repetitions

			injectContentInSdt(endNode, "@EndContext");
			
			

			OpenXmlElement? firstNewElement=null;
			OpenXmlElement? lastNewElement=null;
			for (int i=0;i< data.Count; i++)
			{
				var datai = data[i];

				var contextName = _curRepeatData.ArrayName + "[" + i + "]";

				injectContentInSdt(startNode, "@Context{" + contextName + "}");

				for(int j=0;j<innerContent.Count;j++)
				{
					var newNode=innerContent[j].CloneNode(true);

					commonParent.InsertBefore(newNode, firstElement);

					if (i == 0 && j == 0)
						firstNewElement = newNode;

					lastNewElement = newNode;
				}
			}
			//now strip the inner content
			foreach (var e in innerContent)
				e.Remove();

			if (firstNewElement != null)
			{
				Debug.Assert(lastNewElement != null);
				processNodeRange(firstNewElement, lastNewElement);
			}
			return HandleState.Advance;
		}
		private void injectContentInSdt(SdtElement e,string text)
		{
			if (e is SdtBlock block)
			{
				block.SdtContentBlock = new SdtContentBlock(
				new Paragraph(new Run(new Text(text))));
			}
			else if (e is SdtRun run)
			{
				run.SdtContentRun = new SdtContentRun(
				new Run(new Text(text)));
			}
			else if (e is SdtCell cell)
			{
				cell.SdtContentCell = new SdtContentCell(
					new TableCell(
						new Paragraph(
					new Run(new Text(text)))));
			}
			else
				Debug.Assert(false);
		}
		//private string ConvertToXml(IEnumera)
		private List<OpenXmlElement> getRepeatInnerElements(SdtElement start,SdtElement end,out OpenXmlElement? commonParent)
		{
			var startPath = getXmlPath(start);
			var endPath = getXmlPath(end);

			commonParent = null;
			//get the common parent
			while(startPath.Peek()==endPath.Peek())
			{
				commonParent = startPath.Pop();
				endPath.Pop();
			}
			Debug.Assert(commonParent != null);
			Debug.Assert(startPath.Count > 0);
			Debug.Assert(endPath.Count > 0);

			var startElement = startPath.Peek();
			var endElement = endPath.Peek();

			List<OpenXmlElement> blockElements = new List<OpenXmlElement>();
			bool inside = false;
			foreach(var e in commonParent)
			{
				if (e == startElement)
					inside = true;
				if (inside)
					blockElements.Add(e);
				if (e == endElement)
					break;
			}
			return blockElements;
		}

		private Stack<OpenXmlElement> getXmlPath(OpenXmlElement node)
		{
			Stack<OpenXmlElement> path = new Stack<OpenXmlElement>();
			path.Push(node);
			while (path.Peek().Parent != null)
			{
				path.Push(path.Peek().Parent);
			}
			return path;
		}


		private HandleState handleBindExpr(SdtElement placeholder, string bindExpr, JsonObject context)
		{
			var value = GetBindingExprValue(context, bindExpr)?.ToString();

			if (value != null)
			{
				if (placeholder is SdtBlock block)
				{
					placeholder.Parent?.ReplaceChild(new Paragraph(new Run(
						new Text(value.ToString()))),
						placeholder);
				}
				else if (placeholder is SdtRun)
					placeholder.Parent?.ReplaceChild(new Run(
					new Text(value.ToString())),
					placeholder);
				else if(placeholder is SdtCell)
					placeholder.Parent?.ReplaceChild(
						new SdtCell(
						new SdtContentCell(
					new TableCell(
						new Paragraph(
					new Run(new Text(value.ToString()))))))
						,placeholder);
			}
			return HandleState.Advance;
		}

		private HandleState handleStartContext(SdtElement placeholder, string? newContextName,
			string newContextPath)
		{
			Debug.Assert(placeholder.Parent != null);
			var context = getContext();

			JsonObject? newContext = null; ;
			if (context != null)
				newContext = GetBindingExprValue(context, newContextPath) as JsonObject;

			contextStack.Push(new ContextEntry(newContext, newContextName));

			placeholder.Parent.RemoveChild(placeholder);
			return HandleState.Advance;
		}
		private HandleState handleEndContext(SdtElement placeholder)
		{
			Debug.Assert(placeholder.Parent != null);
			contextStack.Pop();
			placeholder.Parent.RemoveChild(placeholder);
			return HandleState.Advance;
		}


		private JsonNode? GetBindingExprValue(JsonObject baseObj, string bindExpr)
		{
			bindExpr = bindExpr.Trim(' ', '\t');
			var regBind = new Regex(@"^(?<BaseName>(_|\w)(_|\w|\d)*)(\[(?<IndexExpr>\d+)\])?(\.(?<bindExprRest>.+))?$");
			var match = regBind.Match(bindExpr);

			if (!match.Success)
				return null;
			var baseName = match.Groups["BaseName"].Value;

			if (!baseObj.TryGetPropertyValue(baseName, out var baseNode))
				return null;

			var indexExpr = match.Groups["IndexExpr"];
			//if there is an index, evaluate it
			if (indexExpr.Success)
			{
				var indexStr = indexExpr.Value;
				var indexValue = int.Parse(indexStr);
				if (!(baseNode is JsonArray baseArray))
					return null;
				if (indexValue >= baseArray.Count)
					return null;
				baseNode = baseArray[indexValue];
			}
			var rest = match.Groups["bindExprRest"];
			if (rest.Success)
			{
				if (baseNode is JsonObject subObj)
					return GetBindingExprValue(subObj, rest.Value);
				else
					return null;
			}
			return baseNode;
		}
		private WordprocessingDocument _template;

		private Stack<ContextEntry> contextStack = new Stack<ContextEntry>();

		JsonObject? getContext()
		{
			if (contextStack.Count > 0)
				return contextStack.Peek().Context;
			else
				return null;
		}
		JsonObject? getContext(string name)
		{
			foreach (var c in contextStack)
			{
				if (c.Name == name)
					return c.Context;
			}
			return null;
		}
	}

	internal class ContextEntry
	{
		public ContextEntry(JsonObject contextObj, string? name)
		{
			Context = contextObj;
			Name = name;
		}
		public JsonObject Context { private set; get; }
		public string? Name;
	}

	internal class RepeatData
	{
		public RepeatData(SdtElement startPlaceholder,string arrayName,JsonArray data)
		{
			ArrayName = arrayName;
			Data = data;
			StartPlaceholder = startPlaceholder;
		}
		public string ArrayName { private set; get; }
		public JsonArray Data { private set; get; }
		public SdtElement StartPlaceholder { private set; get; }
	}
}