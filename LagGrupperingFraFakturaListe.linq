<Query Kind="Program" />

void Main()
{
	var downloadFolder = @"C:\Users\solvi\Downloads";
	var workFolder = @"C:\MyLocal\Solviken";
	var swExportFolder = Path.Combine(workFolder, "FraStyreweb");
	var swFakturaExportFilNavn = "FakturaOversikt";
	var swFakturaExportFil = Path.Combine(swExportFolder, $"{swFakturaExportFilNavn}.csv");

	CopyNewerFile(Path.Combine(downloadFolder, $"{swFakturaExportFilNavn}.xlsx"), swExportFolder);
	ConvertFromXlsx2Csv(Path.Combine(swExportFolder, $"{swFakturaExportFilNavn}.xlsx"));

	using (var reader = new StreamReader(swFakturaExportFil, Encoding.GetEncoding("UTF-8")))
	{
		reader.ReadLine();      // Skip header
		reader.ReadLine();      // Skip header
		reader.ReadLine();      // Skip header
		string line;
		Console.WriteLine("Visningsnavn");
		while ((line = reader.ReadLine()) != null)
		{
			var fields = line.Split('\t');
			if (fields[0] == string.Empty)
			{
				break;
			}

			var navn = $"{fields[10]} {fields[9]}";
			Console.WriteLine(navn);
		}
	}
}

void ConvertFromXlsx2Csv(string excelFile)
{
	var folder = Path.GetDirectoryName(excelFile);
	var csvFile = Path.Combine(folder, Path.GetFileNameWithoutExtension(excelFile)) + ".csv";

	if (File.GetLastWriteTime(excelFile) > File.GetLastWriteTime(csvFile))
	{
		string scriptName = @"C:\MyLocal\Solviken\xlsx2csv.vbs"; // full path to script
		ProcessStartInfo ps = new ProcessStartInfo();
		ps.FileName = "cscript.exe";
		ps.Arguments = $"{scriptName} {excelFile} {csvFile}";
		ps.WindowStyle = ProcessWindowStyle.Hidden;
		ps.CreateNoWindow = true;
		var process = Process.Start(ps);
		process.WaitForExit();
		process.Close();
	}
}

void CopyNewerFile(string source, string destination)
{
	var fileName = Path.GetFileName(source);
	var destinationFile = Path.Combine(destination, fileName);

	if (File.Exists(source))
	{
		File.Delete(destinationFile);
		File.Move(source, destinationFile);
		Console.WriteLine($"Oppdaterte StyreWeb export fil \"{fileName}\" fra Nedlastinger");
	}
}
