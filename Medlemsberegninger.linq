<Query Kind="Program" />

void Main()
{
	var hwFolder = @"C:\MyLocal\Solviken\FraHavneWeb";
	var swFolder = @"C:\MyLocal\Solviken\FraStyreWeb";
	var hwEksportFil = Path.Combine(hwFolder, "brukereksport_311224.xlsx");
	var hwCsvFil = Path.Combine(hwFolder, "brukereksport_311224.csv");
	var swEksportFil = Path.Combine(swFolder, "Standard_Rapport.csv");
	var styrewebMedlemmer = new HashSet<string>();
	
	using (var reader = new StreamReader(swEksportFil, Encoding.GetEncoding("UTF-8")))
	{
		reader.ReadLine();      // Skip header
		string line;
		while ((line = reader.ReadLine()) != null)
		{
			var fields = line.Split('\t');
			if (fields.Length >= 2)
			{
				var name = $"{fields[1]} {fields[0]}";
				styrewebMedlemmer.Add(name);
				//Console.WriteLine(name);
			}
		}
	}

	ConvertFromXlsx2Csv(hwEksportFil);
	using (var reader = new StreamReader(hwCsvFil, Encoding.GetEncoding("ISO-8859-1")))
	{
		reader.ReadLine();      // Skip header
		string line;
		Console.WriteLine("Godkjente kranbrukere:");
		int antall = 0;
		int gyldige = 0;
		int ugyldige = 0;
		while ((line = reader.ReadLine()) != null)
		{
			var fields = line.Split('\t');
			if (fields.Length >= 49 && fields[3] != "" && fields[49].Contains("Godkjente kranbrukere"))
			{
				antall++;
				var navn = NavnExcel2StyreWeb(fields[3].Trim());
				if (styrewebMedlemmer.Contains(navn))
				{
					gyldige++;
					//Console.WriteLine(navn);
				}
				else
				{
					ugyldige++;
					Console.WriteLine($"*** Ikke medlem lenger: {navn}");
				}
			}
		}

		Console.WriteLine($"Antall m/krankurs: {antall}, nåværende medlemmer: {gyldige}, sluttet: {ugyldige}");
	}
}

private void ConvertFromXlsx2Csv(string excelFile)
{
	var folder = Path.GetDirectoryName(excelFile);
	var csvFile = Path.Combine(folder, Path.GetFileNameWithoutExtension(excelFile)) + ".csv";

	if (!File.Exists(csvFile) || File.GetLastWriteTime(excelFile) > File.GetLastWriteTime(csvFile))
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

string NavnExcel2StyreWeb(string navn)
{
	switch (navn)
	{
		case "Hilde Risan / Christian Schønfeldt":
			return "Hilde Risan";
		case "Simen T. Aasheim":
			return "Simen Aasheim";
		case "Jan Robert  Andersen":
			return "Jan Robert Andersen";
		case "Øyvind Tellefsen (reservert)":
			return "Øyvind Tellefsen";
		case "Joakim Haugen":
			return "Joakim Mordt Haugen";
		case "Joackim H  Hansen":
			return "Joackim H. Hansen";
		case "rune kr.  Stålstrøm":
			return "Rune Stålstrøm";
		case "Harald Olsen (Æ)":
			return "Harald Olsen";
		case "Øystein Olsen (Æ)":
			return "Øystein Olsen";
		case "Jan di Leggerini":
			return "Jan Di Leggerini";
		case "Roald Bartholdsen (Æ)":
			return "Roald Bartholdsen";
		case "Bexrud Bil AS":
			return "Bexrud Bil AS Bexrud Bil AS";
		case "Morten  Gregersen":
			return "Morten Gregersen";
		case "Anders L. S. Herlofsen":
			return "Anders Herlofsen";
		case "Gerhard  Bagge":
			return "Gerhard Bagge";
		case "Morten A. Dundas":
			return "Morten Dundas";

		default:
			return navn;
	}
}

