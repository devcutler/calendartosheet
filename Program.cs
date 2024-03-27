using CultureInfo = System.Globalization.CultureInfo;
using CsvHelper;
using Ical.Net;
using Ical.Net.CalendarComponents;
using CommandLine;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Diagnostics;
using System.Collections.Generic;
using System.IO;
using System.Linq;

internal class Program
{
	class Options
	{
		[Value(0, MetaName = "input", HelpText = "Input file to convert.", Required = true)]
		public string Input { get; set; } = "";
		[Value(1, MetaName = "output", HelpText = "File path to output to.", Required = true)]
		public string Output { get; set; } = "";
		[Option('f', "format", Required = false, HelpText = "Output column values.\nUsable values: name, summary, description, start, end, duration, length, comments, status, location, geolocation, organizer, creator, calendar, type, attendees, created, modified, lastmodified, id, uid, uuid, priority, url, link", Separator = ',',
			Default = new[] { "name", "description", "start", "duration", "end", "comments", "status", "location", "organizer" })]
		public IEnumerable<string> Format { get; set; } = new List<string>();
		[Option('x', "xlsx", Required = false, HelpText = "Output XLSX instead of CSV.", Default = false)]
		public bool XLSX { get; set; }
		[Option('h', "header", Required = false, HelpText = "Output a header row with the names of the columns.", Default = true)]
		public bool Header { get; set; }
	}
	internal static int Main(string[] args)
	{
		ParserResult<Options> opts = Parser.Default.ParseArguments<Options>(args);

		opts.WithParsed(opts =>
		{
			string file = File.ReadAllText(opts.Input);

			Calendar cal = Calendar.Load(file);
			cal.AddTimeZone("America/New_York");

			if (opts.XLSX)
			{
				WriteXLSX(opts, cal);
			}
			else
			{
				WriteCSV(opts, cal);
			}
		});

		return 0;
	}

	private static void WriteXLSX(Options opts, Calendar calendar)
	{
		var wb = new XLWorkbook();
		var ws = wb.Worksheets.Add("Events");

		int row = 1;

		if (opts.Header)
		{
			int col = 1;
			foreach (string key in opts.Format)
			{
				ws.Cell(row, col).SetValue(key);
				col++;
			}
			row++;
		}

		foreach (CalendarEvent item in calendar.Events)
		{
			int col = 1;
			foreach (string key in opts.Format)
			{
				ws.Cell(row, col).SetValue(GetEventDetail(item, key));
				col++;
			}
			row++;
		}
		wb.SaveAs(opts.Output);
	}

	static void WriteCSV(Options opts, Calendar calendar)
	{
		using var writer = new StreamWriter(opts.Output);
		using var csv = new CsvWriter(writer, CultureInfo.InvariantCulture);

		if (opts.Header)
		{
			foreach (string key in opts.Format)
			{
				csv.WriteField(key);
			}
			csv.NextRecord();
		}

		foreach (CalendarEvent item in calendar.Events)
		{
			foreach (string key in opts.Format)
			{
				csv.WriteField(GetEventDetail(item, key));
			}

			csv.NextRecord();
		}
	}
	static string GetEventDetail(CalendarEvent item, string key)
	{
		return key.ToLower() switch
		{
			"name" or "summary" => item?.Summary ?? "",
			"description" => item?.Description ?? "",
			"start" => item?.Start.ToString() ?? "",
			"end" => item?.End?.ToString() ?? "",
			"duration" or "length" => item?.Duration.ToString() ?? "",
			"comments" => string.Join(',', item.Comments),
			"status" => item?.Status ?? "",
			"location" => item?.Location ?? "",
			"geolocation" => item.GeographicLocation.ToString(),
			"organizer" or "creator" => item?.Organizer?.CommonName?.ToString() ?? "",
			"calendar" => item.Calendar.Name,
			"type" => item.Name,
			"attendees" => string.Join(',', item.Attendees),
			"created" => item.Created.ToString(),
			"modified" or "lastmodified" => item.LastModified.ToString(),
			"id" or "uid" or "uuid" => item.Uid,
			"priority" => item.Priority.ToString(),
			"url" or "link" => item.Url.ToString(),
			_ => ""
		};
	}
}