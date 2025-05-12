<Query Kind="Program" />

void Main()
{
	//VisAlledata(new StyreWebExport().LesData("1"));
	//VisAlledata(new ExcelExport().LesData());
	//VisAlledata(new HavneWebExport().LesData());
	
	VisEierEndringer(new HavneWebExport().LesData(), new StyreWebExport().LesData());
}

void VisEierEndringer(HavneData havn1, HavneData havn2)
{
	var plasser1 = havn1.GetAllePlasser();
	var plasser2 = havn2.GetAllePlasser();
	var tapere = new List<(string, string)>();      // Eier, plassId
	var vinnere = new List<(string, string)>();      // Eier, plassId

	Console.WriteLine($"Båtplass eier endringer fra {havn1.Navn} til {havn2.Navn}:\n");
	
	foreach (var plass1 in plasser1)
	{
		var plass2 = havn2.GetBatPlass(plass1.PlassId);
		if (plass2 != null)
		{
			if (plass2.Eier != plass1.Eier)
			{
				var taper = plass1.Eier ?? "Ledig";
				var vinner = plass2.Eier ?? "Ledig";
				tapere.Add((taper, plass1.PlassId));
				vinnere.Add((vinner, plass2.PlassId));
				
				//Console.WriteLine($"{plass1.PlassId}: Eier fra {taper} til {vinner}");
			}
		}
		else
		{
			Console.WriteLine($"{plass1.PlassId} mangler i {havn2.Navn}");
		}
	}
	
	Console.WriteLine("\nInnskudd som skal krediteres:");
	var venteListe = new List<string>();
	foreach (var taper in tapere)
	{
		if (!vinnere.Any(v => v.Item1 == taper.Item1))
		{
			// Gitt fra seg plass uten å få ny
			var plass = havn1.GetBatPlass(taper.Item2);
			var innskudd = plass.Innskudd.ToString("n").PadLeft(9);
			var plass2025 = havn2.GetBatPlass(taper.Item2);
			if (plass2025.Eier != null)
			{
				Console.WriteLine($"{taper.Item2}: kr. {innskudd} - {taper.Item1}");
			}
			else
			{
				venteListe.Add($"{taper.Item2}: kr. {innskudd} - {taper.Item1}");
			}
		}
	}
	
	Console.WriteLine("\nInnskudd som krediteres etter at plassen er videresolgt:");
	foreach (var vente in venteListe)
	{
		Console.WriteLine(vente);
	}

	Console.WriteLine("\nNye innskudd:");
	foreach (var vinner in vinnere)
	{
		if (!tapere.Any(t => t.Item1 == vinner.Item1))
		{
			// Fått ny plass uten å gi fra seg en
			var innskudd = havn2.BeregnInnskudd(vinner.Item2).ToString("n").PadLeft(9);
			Console.WriteLine($"{vinner.Item2}: kr. {innskudd} - {vinner.Item1}");
		}
	}
}

void VisAlledata(HavneData dataSet)
{
	var andelsplasser = dataSet.GetAndelsPlasser();
	var sesongplasser = dataSet.GetSesongPlasser();
	var ungdomsplasser = dataSet.GetUngdomsPlasser();
	var ledigeplasser = dataSet.GetLedigePlasser();

	Console.Write($"Eksport fra {dataSet.Navn}");
	if (dataSet.PlassPrefix != null)
	{
		Console.WriteLine($" - Plasser som starter med \"{dataSet.PlassPrefix}\"");
	}
	else
	{
		Console.WriteLine();
	}
	
	Console.WriteLine($"\n{andelsplasser.Count} andelsplasser");
	foreach (var plass in andelsplasser)
	{
		var bredde = ((double)plass.BatBredde / 100).ToString("0.00");
		var lengde = ((double)plass.BatLengde / 100).ToString("0.00");
		Console.WriteLine($"{plass.PlassId}: {plass.Eier}, BxL={bredde}, {lengde}");
	}

	Console.WriteLine($"\n{sesongplasser.Count} sesongplasser");
	foreach (var plass in sesongplasser)
	{
		var eier = plass.Eier != null ? $" (fra {plass.Eier})" : "";
		Console.WriteLine($"{plass.PlassId}: {plass.Leier}{eier}");
	}

	Console.WriteLine($"\n{ungdomsplasser.Count} ungdomsplasser");
	foreach (var plass in ungdomsplasser)
	{
		Console.WriteLine($"{plass.PlassId}: {plass.Leier}");
	}

	Console.WriteLine($"\n{ledigeplasser.Count} ledige plasser");
	foreach (var plass in ledigeplasser)
	{
		Console.WriteLine($"{plass.PlassId}");
	}
}

public class BatPlass
{
	public string PlassId { get; set; }
	public string Eier { get; set; }
	public string Leier { get; set; }
	public int BatBredde { get; set; }
	public int BatLengde { get; set; }
	public int LysApning { get; set; }
	public int Innskudd { get; set; }
	public bool UngdomsPlass { get;set; }
}

public abstract class HavneData
{
	protected abstract SortedDictionary<string, BatPlass> BatPlasser { get; set; }
	protected abstract HavneData Read();
	
	public abstract string Navn { get; }
	
	public string PlassPrefix { get; set; }

	public HavneData LesData(string prefix = null)
	{
		PlassPrefix = prefix;
		Read();
		
		if (prefix != null)
		{
			var excluded = BatPlasser.Keys.Where(p => !p.StartsWith(prefix)).ToList();
			foreach (var plassId in excluded)
			{
				BatPlasser.Remove(plassId);
			}
		}
		
		return this;
	}
	
	public BatPlass GetBatPlass(string plassId)
	{
		if (BatPlasser.TryGetValue(plassId, out var plass))
		{
			return plass;
		}

		return null;
	}
	
	public List<BatPlass> GetAllePlasser()
	{
		return BatPlasser.Values.ToList();
	}
	
	public List<BatPlass> GetAndelsPlasser()
	{
		return BatPlasser.Values.Where(v => v.Eier != null).ToList();
	}
	
	public List<BatPlass> GetSesongPlasser()
	{
		return BatPlasser.Values.Where(v => v.Leier != null && !v.UngdomsPlass).ToList();
	}

	public List<BatPlass> GetUngdomsPlasser()
	{
		return BatPlasser.Values.Where(v => v.UngdomsPlass).ToList();
	}

	public List<BatPlass> GetLedigePlasser()
	{
		return BatPlasser.Values
		.Except(GetAndelsPlasser())
		.Except(GetSesongPlasser())
		.Except(GetUngdomsPlasser())
		.ToList();
	}
	
	protected string NavnExcel2StyreWeb(string navn)
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

			default:
				return navn;
		}
	}
	
	public int BeregnInnskudd(string plassId)
	{
		var plass = GetBatPlass(plassId);
		if (plass != null)
		{
			double lengde = (double)plass.BatLengde / 100;
			switch (lengde)
			{
				case double len when (len <= 5.4):
					return 7750;
				case double len when (len <= 7.0):
					return 11100;
				case double len when (len <= 8.7):
					return 14700;
				case double len when (len <= 9.1):
					return 19750;
				case double len when (len <= 10.0):
					return 22500;
				case double len when (len <= 10.6):
					return 25150;
				case double len when (len <= 11.8):
					return 27900;
				case double len when (len <= 12.4):
					return 34500;
				default:
					return 45500;
			}
		}
		
		return -1;
	}
}

public class StyreWebExport : HavneData
{
	private string swMarinaFil;
	private string swFramleieFil;

	protected override SortedDictionary<string, BatPlass> BatPlasser { get; set; }
	public override string Navn { get => "StyreWeb"; }

	public StyreWebExport()
	{
		var workFolder = @"C:\MyLocal\Solviken";
		var swExportFolder = Path.Combine(workFolder, "FraStyreweb");
		swMarinaFil = Path.Combine(swExportFolder, "Marina.csv");
		swFramleieFil = Path.Combine(swExportFolder, "Fremleie_historie.csv");
		BatPlasser = new SortedDictionary<string, BatPlass>();
	}

	protected override HavneData Read()
	{
		CopyNewerFile(@"C:\Users\solvi\Downloads\Marina.csv", @"C:\MyLocal\Solviken\FraStyreweb");
		CopyNewerFile(@"C:\Users\solvi\Downloads\Fremleie_historie.csv", @"C:\MyLocal\Solviken\FraStyreweb");

		using (var reader = new StreamReader(swMarinaFil, Encoding.GetEncoding("UTF-8")))
		{
			reader.ReadLine();      // Skip header
			string line;
			while ((line = reader.ReadLine()) != null && !line.StartsWith("#"))
			{
				var fields = line.Split('\t');
				var plassId = fields[1];
				var plassType = fields[2];
				var eier = fields[8];
				var breddeMeter = fields[3];
				var lengdeMeter = fields[4];
				int breddeCm = 0;
				int lengdeCm = 0;
				
				if (eier == "Solviken Båtforening" || eier == "" || eier == "Ledig")
				{
					eier = null;
				}
				
				if (double.TryParse(breddeMeter, out var bredde))
				{
					breddeCm = (int)Math.Round(bredde * 100);
				}
				
				if (double.TryParse(lengdeMeter, out var lengde))
				{
					lengdeCm = (int)Math.Round(lengde * 100);
				}

				var ungdomsPlass = (plassType == "Ungdomsplass");
				BatPlasser[plassId] = new BatPlass
				{
					PlassId = plassId,
					Eier = eier,
					BatBredde = breddeCm,
					BatLengde = lengdeCm,
					UngdomsPlass = ungdomsPlass
				};
			}
		}
		
		using (var reader = new StreamReader(swFramleieFil, Encoding.GetEncoding("UTF-8")))
		{
			reader.ReadLine();      // Skip header
			string line;
			while ((line= reader.ReadLine()) != null && !line.StartsWith("#"))
			{
				var fields = line.Split('\t');
				var plassId = fields[0];
				var tilDato = fields[3];
				var leier = fields[4];
				
				if (leier == "")
				{
					leier = null;
				}

				if (BatPlasser.TryGetValue(plassId, out var batPlass))
				{
					if (DateTime.TryParse(tilDato, out var sluttDato))
					{
						batPlass.Leier = sluttDato > DateTime.Now ? leier : null;
					}
					else
					{
						leier = null;
					}
				}
			}
		}
		
		return this;
	}

	private void CopyNewerFile(string source, string destination)
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
}

public class ExcelExport : HavneData
{
	private string batplassFil;

	protected override SortedDictionary<string, BatPlass> BatPlasser { get; set; }
	public override string Navn { get => "Excel"; }

	public ExcelExport()
	{
		var workFolder = @"C:\MyLocal\Solviken";
		batplassFil = Path.Combine(workFolder, "Batplasser.csv");
		BatPlasser = new SortedDictionary<string, BatPlass>();
	}

	protected override HavneData Read()
	{
		using (var reader = new StreamReader(batplassFil, Encoding.GetEncoding("ISO-8859-1")))
		{
			reader.ReadLine();      // Skip header
			string line;
			while ((line = reader.ReadLine()) != null)
			{
				var fields = line.Split('\t');
				var plassId = fields[2].Split(' ')[0].Trim('"');
				var eier = NavnExcel2StyreWeb(fields[11].Trim());
				var leier = fields[12].Trim();
				var breddeMeter = fields[4];
				var lengdeMeter = fields[5];
				int breddeCm = 0;
				int lengdeCm = 0;

				if (eier == "Solviken Båtforening" || eier == "" || eier == "Ledig")
				{
					eier = null;
				}
				
				if (leier == "")
				{
					leier = null;
				}

				int.TryParse(breddeMeter, out breddeCm);
				int.TryParse(lengdeMeter, out lengdeCm);

				BatPlasser[plassId] = new BatPlass
				{
					PlassId = plassId,
					Eier = eier,
					Leier = leier,
					BatBredde = breddeCm,
					BatLengde = lengdeCm
				};
			}
		}

		return this;
	}
}

public class HavneWebExport : HavneData
{
	private string hwExportFil;

	protected override SortedDictionary<string, BatPlass> BatPlasser { get; set; }
	public override string Navn { get => "HavneWeb"; }

	public HavneWebExport()
	{
		var workFolder = @"C:\Users\solvi\OneDrive\Solviken\2025\Havnedatabasen";
		hwExportFil = Path.Combine(workFolder, "SolvikenBtforening_311224_124226.csv");
		BatPlasser = new SortedDictionary<string, BatPlass>();
	}

	protected override HavneData Read()
	{
		using (var stream = new FileStream(hwExportFil, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
		{
			using (var reader = new StreamReader(stream, Encoding.GetEncoding("ISO-8859-1")))
			{
				reader.ReadLine();      // Skip header
				string line;
				while ((line = reader.ReadLine()) != null)
				{
					var fields = line.Split('\t');
					var plassId = fields[2].Split(' ')[0];
					var eier = NavnExcel2StyreWeb(fields[14].Trim());
					var leier = fields[21].Trim();
					var innskudd = fields[9];
					var breddeMeter = fields[4];
					var lengdeMeter = fields[5];
					int breddeCm = 0;
					int lengdeCm = 0;
					int innskuddKr = 0;

					if (eier == "Solviken Båtforening" || eier == "" || eier == "Ledig")
					{
						eier = null;
					}

					if (leier == "")
					{
						leier = null;
					}

					int.TryParse(breddeMeter, out breddeCm);
					int.TryParse(lengdeMeter, out lengdeCm);
					int.TryParse(innskudd, out innskuddKr);

					BatPlasser[plassId] = new BatPlass
					{
						PlassId = plassId,
						Eier = eier,
						Innskudd = innskuddKr,
						Leier = leier,
						BatBredde = breddeCm,
						BatLengde = lengdeCm
					};
				}
			}
		}

		return this;
	}
}