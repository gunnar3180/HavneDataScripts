<Query Kind="Program">
  <Reference>&lt;RuntimeDirectory&gt;\System.Collections.Concurrent.dll</Reference>
  <Namespace>System.Collections.Concurrent</Namespace>
</Query>

void Main()
{
	VisAlledata(new StyreWebExport().LesData("1"));
	//VisAlledata(new ExcelExport().LesData());
	//VisAlledata(new HavneWebExport().LesData());

	//VisEierEndringer(new HavneWebExport().LesData(), new StyreWebExport().LesData());
	//VisEierEndringer(new ExcelExport().LesData(), new StyreWebExport().LesData());
	//new List<int>{1, 2, 3, 5, 6}.ForEach(x => VisArealForskjeller(new HavneWebExport().LesData(x.ToString()), new StyreWebExport().LesData(x.ToString())));
	//VisArealForskjeller(new HavneWebExport().LesData("6"), new StyreWebExport().LesData("6"));
	//VisVaktFritak(new HavneWebExport().LesData());
	//BeregnBatplassAvgifter(new StyreWebExport().LesData());
	//VisAlleMedVaktplikt(new StyreWebExport().LesData(), new HavneWebExport().LesData());
	//VisLedigePlasser(new StyreWebExport().LesData());
	//VisAlleMedVaktfritakOgPlasser(new StyreWebExport().LesData());
	//SjekkSesongOgUngdom(new StyreWebExport().LesData());
	//FinnLeietillegg(new StyreWebExport().LesData());
}

void FinnLeietillegg(HavneData havn)
{
	var leietilleggGrupper = new ConcurrentDictionary<string, List<(string, string)>>();	// Bruker, plassId
	var sesongPlasser = havn.GetSesongPlasser();
	foreach (var plass in sesongPlasser)
	{
		var gruppe = havn.BeregnInnskudd(plass.PlassId).Item2;
		var gruppeBrukere = leietilleggGrupper.GetOrAdd(gruppe, a => new List<(string, string)>());
		gruppeBrukere.Add((plass.Leier, plass.PlassId));
	}
	
	foreach (var gruppe in leietilleggGrupper.OrderBy(g => g.Key))
	{
		if (gruppe.Value.Count > 0)
		{
			var leietillegg = LeietilleggFraGruppe(gruppe.Key);
			Console.WriteLine($"\nGruppe {gruppe.Key} ({gruppe.Value.Count} stk, kr. {leietillegg}):");
			foreach (var bruker in gruppe.Value)
			{
				Console.WriteLine($"{bruker.Item1} ({bruker.Item2})");
			}
		}
	}
}

int LeietilleggFraGruppe(string gruppe)
{
	switch (gruppe)
	{
		case "A":
		case "B":
			return 1900;
		case "C":
			return 2000;
		case "D":
			return 3000;
		default:
			return 4000;
	}
}

void SjekkSesongOgUngdom(HavneData havn)
{
	var sesongOgUngdom = havn.GetSesongPlasser().Concat(havn.GetUngdomsPlasser()).OrderBy(h => h.PlassId);
	var framleiePlasser = havn.GetAllePlasser().Where(h => h.Leier != null).OrderBy(h => h.PlassId);
	if (sesongOgUngdom.SequenceEqual(framleiePlasser))
	{
		Console.WriteLine("Alle framleieplasser registrert som enten sesong eller ungdomsplass");
	}
	else
	{
		var feil1 = sesongOgUngdom.Except(framleiePlasser);
		var feil2 = framleiePlasser.Except(sesongOgUngdom);
		
		if (feil1.Count() > 0)
		{
			Console.WriteLine("Følgende sesong/ungdom ikke registrert som framleie:");
			foreach (var feil in feil1)
			{
				Console.WriteLine(feil.PlassId);
			}
		}
		if (feil2.Count() > 0)
		{
			Console.WriteLine("Følgende framleieplasser ikke registrert som sesong/ungdom:");
			foreach (var feil in feil2)
			{
				Console.WriteLine(feil.PlassId);
			}
		}
	}
}

void VisAlleMedVaktfritakOgPlasser(HavneData havn)
{
	var fritaksPlasser = havn.GetAllePlasser().Where(p => p.Vaktfritak != null);
	var sortert = fritaksPlasser.OrderBy(p => p.Leier??p.Eier);
	var gruppert = sortert.GroupBy(s => s.Leier??s.Eier);
	var sortertPaArsak = new Dictionary<string, List<IGrouping<string, BatPlass>>>();		// Årsak, brukere
	Console.WriteLine("Medlemmer med vaktfritak:");
	int antallFritak = 0;
	foreach (var bruker in gruppert)
	{
		var forstePlass = bruker.ElementAt(0);
		var vaktFritak = forstePlass.Vaktfritak;
		if (sortertPaArsak.TryGetValue(vaktFritak, out var brukere))
		{
			brukere.Add(bruker);
		}
		else
		{
			brukere = new List<IGrouping<string, BatPlass>>();
			brukere.Add(bruker);
			sortertPaArsak[vaktFritak] = brukere;
		}
	}

	foreach (var arsak in sortertPaArsak)
	{
		Console.WriteLine($"\n--- {arsak.Key} ---");
		int arsakPlasser = 0;
		foreach (var bruker in arsak.Value)
		{
			Console.Write($"{bruker.Key,-30}");
			int plasser = 0;
			foreach (var plass in bruker)
			{
				var separator = plasser == 0 ? "" : ", ";
				Console.Write($"{separator}{plass.PlassId}");
				plasser++;
				antallFritak++;
				arsakPlasser++;
			}
			Console.WriteLine();
		}

		Console.WriteLine($"-Antall fritaksplasser: {arsakPlasser}");
	}
	
	Console.WriteLine($"\nAntall medlemmer med vaktfritak: {gruppert.Count()}, båtplasser: {antallFritak}");
}

void VisLedigePlasser(HavneData havn)
{
	var ledige = havn.GetLedigePlasser();
	var sortert = ledige.OrderBy(l => l.LysApning);
	foreach (var plass in sortert)
	{
		var bredde = plass.LysApning > 0 ? (double)plass.LysApning / 100 : 0;
		Console.WriteLine($"{plass.PlassId}: {bredde} m");
	}
}

void VisAlleMedVaktplikt(HavneData styreWeb, HavneData hwExport)
{
	// Lag liste for import til gruppering i styreweb
	var pliktigePlasser = styreWeb.GetAndelsPlasser().Concat(styreWeb.GetSesongPlasser());
	var fritak2024 = hwExport.GetAllePlasser().Where(p => p.Vaktfritak != null);
	Console.WriteLine($"Visningsnavn");
	foreach (var plass in pliktigePlasser)
	{
		Console.Write($"{plass.Leier ?? plass.Eier};{plass.PlassId}");
		if (fritak2024.Any(f => f.PlassId == plass.PlassId))
		{
			Console.WriteLine($";Fritak 2024");
		}
		else
		{
			Console.WriteLine();
		}
	}
}

void BeregnBatplassAvgifter(HavneData havneData)
{
	int totalFakturering = 0;
	foreach (var plass in havneData.GetAllePlasser().Except(havneData.GetLedigePlasser()))
	{
		var batplassAvgift = plass.BeregnBatplassAvgift();
		totalFakturering += batplassAvgift;
		Console.WriteLine($"{plass.PlassId}: {(plass.Leier ?? plass.Eier),-30} (BxL: {plass.BatBredde}x{plass.BatLengde}) kr. {batplassAvgift:n}");
	}

	Console.WriteLine($"Total fakturering av båtplassavgifter i 2025: kr. {totalFakturering:n}");
}

void VisVaktFritak(HavneData havneData)
{
	Console.WriteLine("Vaktfritak i 2024:\n");
	int antall = 0;
	foreach (var plass in havneData.GetAllePlasser())
	{
		if ((plass.Eier != null || plass.Leier != null) && plass.Vaktfritak != null)
		{
			antall++;
			Console.WriteLine($"{plass.PlassId}: {plass.Leier ?? plass.Eier}");
		}
	} 

	Console.WriteLine($"\nTotalt {antall} vaktfritak");
}

void VisArealForskjeller(HavneData havn1, HavneData havn2)
{
	var plasser1 = havn1.GetAllePlasser();
	var plasser2 = havn2.GetAllePlasser();
	
	if (havn1.PlassPrefix != null)
	{
		Console.WriteLine($"\nBåtplasser som starter med {havn1.PlassPrefix}:");
	}
	else
	{
		Console.WriteLine("\nAlle båtplasser:");
	}
	
	Console.WriteLine($"Båtplass areal endringer fra {havn1.Navn} til {havn2.Navn} (for plasser med samme eier):\n");

	Console.WriteLine($"Plass Bruker                   {havn1.Navn}      -  {havn2.Navn}");
	Console.WriteLine($"----------------------------------------------------------------");
	foreach (var plass1 in plasser1)
	{
		var plass2 = havn2.GetBatPlass(plass1.PlassId);
		if (plass2 != null)
		{
			var bruker1 = plass1.Leier ?? plass1.Eier;
			var bruker2 = plass2.Leier ?? plass2.Eier;
			if (bruker1 == bruker2 && bruker1 != null)
			{
				if (plass1.BatBredde.Ulik(plass2.BatBredde) || plass1.BatLengde.Ulik(plass2.BatLengde))
				{
					var bredde1 = ((double)plass1.BatBredde / 100).ToString("0.00");
					var lengde1 = ((double)plass1.BatLengde / 100).ToString("0.00");
					var bredde2 = ((double)plass2.BatBredde / 100).ToString("0.00");
					var lengde2 = ((double)plass2.BatLengde / 100).ToString("0.00");
					Console.WriteLine($"\n{plass1.PlassId}: {bruker1,-20}BxL ({bredde1} x {lengde1}) - ({bredde2} x {lengde2}) - {plass1.batType}");
				}
			}
		}
		else
		{
			Console.WriteLine($"{plass1.PlassId} mangler i {havn2.Navn}");
		}
	}
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
				
				Console.WriteLine($"{plass1.PlassId}: Eier fra {taper} til {vinner}");
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
			var innskudd = havn2.BeregnInnskudd(vinner.Item2).Item1.ToString("n").PadLeft(9);
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
	var medlemsRegister = new MedlemsRegister().LesData();
	
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
		PrintBatplass(plass, true, medlemsRegister);
	}

	Console.WriteLine($"\n{sesongplasser.Count} sesongplasser");
	foreach (var plass in sesongplasser)
	{
		var eier = plass.Eier != null ? $"(fra {plass.Eier})" : null;
		PrintBatplass(plass, false, medlemsRegister, eier);
	}

	Console.WriteLine($"\n{ungdomsplasser.Count} ungdomsplasser");
	foreach (var plass in ungdomsplasser)
	{
		var tlf = medlemsRegister.Medlemmer[plass.Leier].Tlf;
		Console.WriteLine($"{plass.PlassId}: {plass.Leier,-25}  {tlf,-13}");
	}

	Console.WriteLine($"\n{ledigeplasser.Count} ledige plasser");
	var lysListe = new List<(string, string)>();
	
	foreach (var plass in ledigeplasser)
	{
		Console.Write($"{plass.PlassId}");
		int lysApning = plass.LysApning;
		if (lysApning > 0)
		{
			var lysApningMeter = ((double)lysApning / 100).ToString("0.00");
			Console.WriteLine($": {lysApningMeter} m");
			lysListe.Add((plass.PlassId, lysApningMeter));
		}
		else
		{
			Console.WriteLine();
		}
	}
	
	var sortert = lysListe.OrderBy(l => l.Item2);
	
	Console.WriteLine("\nLedige plasser sortert på lysåpning:");
	foreach (var plass in sortert)
	{
		Console.WriteLine($"{plass.Item1}: {plass.Item2} m");
	}
}

void PrintBatplass(BatPlass plass, bool visEier, MedlemsRegister medlemsRegister, string postfix = null)
{
	var bruker = visEier ? plass.Eier : plass.Leier;
	var bredde = ((double)plass.BatBredde / 100).ToString("0.00");
	var lengde = ((double)plass.BatLengde / 100).ToString("00.00").TrimStart('0').PadLeft(5);
	var tlf = medlemsRegister.Medlemmer[bruker].Tlf;
	var eier = TilpassNavn(bruker);
	var lysApning = plass.LysApning > 0 ? ((double)plass.LysApning / 100).ToString("0.00") : "---";
	Console.Write($"{plass.PlassId.Substring(0, 4)}: {eier,-25} {tlf,-13} BxL: {bredde} x {lengde}  LÅ: {lysApning}");
	if (postfix != null)
	{
		Console.WriteLine($"  {postfix}");
	}
	else
	{
		Console.WriteLine();
	}
}

string TilpassNavn(string navn)
{
	switch (navn)
	{
		case "Tore Gulbrandsen Schjelderup":
			return "Tore G. Schjelderup";
		case "Bexrud Bil AS Bexrud Bil AS":
			return "Bexrud Bil AS";
		default:
			return navn;
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
	public bool SesongPlass { get; set; }
	public bool UngdomsPlass { get; set; }
	public bool Reservert { get; set; }
	public string Vaktfritak { get; set; }
	public string batType { get; set; }

	public int BeregnBatplassAvgift()
	{
		if (UngdomsPlass)
		{
			return 1000;
		}
		
		double bredde = (double)BatBredde / 100;
		double lengde = (double)BatLengde / 100;
		int beregnetAvgift = (int)Math.Round(bredde * lengde * PrisFaktor(lengde));
		if (beregnetAvgift < 2500)
		{
			beregnetAvgift = 2500;
		}

		var leiePlass = Leier != null;
		return beregnetAvgift + (leiePlass ? LeieTillegg(lengde) : 0);
	}

	int PrisFaktor(double lengde)
	{
		return lengde <= 7.2 ? 160 : (lengde >= 9.2 ? 200 : 180);
	}

	int LeieTillegg(double lengde)
	{
		switch (lengde)
		{
			case double len when (len <= 7.0):
				return 1900;
			case double len when (len > 7.0 && len <= 8.7):
				return 2000;
			case double len when (len > 8.7 && len <= 9.1):
				return 3000;
			default:
				return 4000;
		}
	}
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
		return BatPlasser.Values.Where(v => v.SesongPlass).ToList();
	}

	public List<BatPlass> GetUngdomsPlasser()
	{
		return BatPlasser.Values.Where(v => v.UngdomsPlass).ToList();
	}

	public List<BatPlass> GetReservertePlasser()
	{
		return BatPlasser.Values.Where(v => v.Reservert).ToList();
	}

	public List<BatPlass> GetLedigePlasser()
	{
		return BatPlasser.Values
		.Except(GetAndelsPlasser())
		.Except(GetSesongPlasser())
		.Except(GetUngdomsPlasser())
		.Except(GetReservertePlasser())
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
	
	public (int, string) BeregnInnskudd(string plassId)
	{
		var plass = GetBatPlass(plassId);
		if (plass != null)
		{
			double lengde = (double)plass.BatLengde / 100;
			switch (lengde)
			{
				case double len when (len <= 5.4):
					return (7750, "A");
				case double len when (len <= 7.0):
					return (11100, "B");
				case double len when (len <= 8.7):
					return (14700, "C");
				case double len when (len <= 9.1):
					return (19750, "D");
				case double len when (len <= 10.0):
					return (22500, "E");
				case double len when (len <= 10.6):
					return (25150, "F");
				case double len when (len <= 11.8):
					return (27900, "G");
				case double len when (len <= 12.4):
					return (34500, "H");
				default:
					return (45500, "L");
			}
		}
		
		return (-1, "X");
	}
}

public class StyreWebExport : HavneData
{
	private string downloadFolder;
	private string swMarinaFil;
	private string swFramleieFil;
	private string swGruppeVaktplikt;
	private string swExportFolder;
	private List<string> swFritaksGrupper;

	protected override SortedDictionary<string, BatPlass> BatPlasser { get; set; }
	public override string Navn { get => "StyreWeb"; }

	public StyreWebExport()
	{
		downloadFolder = @"C:\Users\solvi\Downloads";
		var workFolder = @"C:\MyLocal\Solviken";
		swExportFolder = Path.Combine(workFolder, "FraStyreweb");
		swMarinaFil = Path.Combine(swExportFolder, "Marina.csv");
		swFramleieFil = Path.Combine(swExportFolder, "Fremleie_historie.csv");
		swGruppeVaktplikt = "Vaktplikt-2025";
		swFritaksGrupper = new List<string>
		{
			"Styre",
			"Havneutvalget",
			"Revisorer",
			"Valgkomite",
			"Elektrikergruppa",
			"Diverse_verv",
			"Omsøkt_vaktfritak",
		};
		
		BatPlasser = new SortedDictionary<string, BatPlass>();
	}

	protected override HavneData Read()
	{
		CopyNewerFile(Path.Combine(downloadFolder, "Marina.csv"), swExportFolder);
		CopyNewerFile(Path.Combine(downloadFolder, "Fremleie_historie.csv"), swExportFolder);
		//CopyNewerFile(Path.Combine(downloadFolder, "Fremleie_historie.csv"), swExportFolder);
		
		foreach (var gruppe in swFritaksGrupper.Append(swGruppeVaktplikt))
		{
			CopyNewerFile(Path.Combine(downloadFolder, $"Gruppe{gruppe}.xlsx"), swExportFolder);
			ConvertFromXlsx2Csv(Path.Combine(swExportFolder, $"Gruppe{gruppe}.xlsx"));
		}

		var oppmaling = new LysApninger().Read();
		
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
				
				if (plassType == "Kan ikke brukes")
				{
					continue;
				}
				
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

				var sesongPlass = (plassType == "Sesongplass");
				var ungdomsPlass = (plassType == "Ungdomsplass");
				var reservert = (plassType == "Reservert");
				
				BatPlasser[plassId] = new BatPlass
				{
					PlassId = plassId,
					Eier = eier,
					BatBredde = breddeCm,
					BatLengde = lengdeCm,
					SesongPlass = sesongPlass,
					UngdomsPlass = ungdomsPlass,
					Reservert = reservert,
					LysApning = oppmaling.GetLysApning(plassId)
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
				
				if (leier != "")
				{
					if (BatPlasser.TryGetValue(plassId, out var batPlass))
					{
						if (DateTime.TryParse(tilDato, out var sluttDato) &&
							sluttDato > DateTime.Now)
						{
							batPlass.Leier = leier;
							//Console.WriteLine($"{plassId} er leid av {leier}");
						}
					}
				}
			}
		}

		var kjenteFritak = swFritaksGrupper
		.SelectMany(g => LesGruppe(g))
		.OrderBy(g => g.Item1)
		.Distinct(new VervNavnComparer())
		.ToDictionary(key => key.Item1, value => value.Item2);
		var pliktigePlasser = GetAndelsPlasser().Concat(GetSesongPlasser());
		var vaktpliktige = LesGruppe(swGruppeVaktplikt);
		var fritak = pliktigePlasser.Where(p => !vaktpliktige.Any(v => (p.Leier??p.Eier) == v.Item1));
		
		foreach (var plass in fritak)
		{
			if (kjenteFritak.TryGetValue(plass.Leier??plass.Eier, out var reason))
			{
				plass.Vaktfritak = reason;
			}
			else
			{
				plass.Vaktfritak = "Fritak, ukjent årsak";
			}
		}

		return this;
	}

	private List<(string, string)> LesGruppe(string gruppe)
	{
		var gruppeFil = Path.Combine(swExportFolder, $"Gruppe{gruppe}.csv");
		var medlemmer = new List<(string, string)>();	// (Navn, gruppe)
		using (var reader = new StreamReader(gruppeFil, Encoding.GetEncoding("UTF-8")))
		{
			reader.ReadLine();      // Skip header
			reader.ReadLine();      // Skip header
			reader.ReadLine();      // Skip header
			string line;
			while ((line = reader.ReadLine()) != null)
			{
				var fields = line.Split('\t');
				if (fields[0] == string.Empty)
				{
					break;
				}

				var navn = $"{fields[1]} {fields[0]}";
				medlemmer.Add((navn, gruppe));
			}

			return medlemmer;
		}
	}

	private void ConvertFromXlsx2Csv(string excelFile)
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
					var vaktFritak = fields[11] == "on" ? "Fritak" : null;
					var batType = leier != string.Empty ? fields[24] : fields[17];
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
						BatLengde = lengdeCm,
						Vaktfritak = vaktFritak,
						batType = batType
					};
				}
			}
		}

		return this;
	}
}

public class LysApninger
{
	private string lysApningFil;
	private Dictionary<string, int> BatPlasser { get; set; }

	public LysApninger()
	{
		var workFolder = @"C:\Users\solvi\OneDrive\Solviken\2025\Havnedatabasen";
		lysApningFil = Path.Combine(workFolder, "LysApninger.csv");
		BatPlasser = new Dictionary<string, int>();
	}

	public LysApninger Read()
	{
		using (var stream = new FileStream(lysApningFil, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
		{
			using (var reader = new StreamReader(stream, Encoding.GetEncoding("ISO-8859-1")))
			{
				reader.ReadLine();      // Skip header
				string line;
				while ((line = reader.ReadLine()) != null)
				{
					var fields = line.Split(';');
					var plassId = fields[0];
					var lysApningString = fields[1];
					if (double.TryParse(lysApningString, out var lysApning))
					{
						BatPlasser[plassId] = (int)Math.Round(lysApning * 100);	// Unit = cm
					}
				}
			}
		}
		
		return this;
	}
	
	public int GetLysApning(string plassId)
	{
		if (BatPlasser.TryGetValue(plassId, out var lysApning))
		{
			return lysApning;
		}
		
		return -1;
	}
}

public static class Extensions
{
	public static bool Ulik(this int measure1, int measure2)
	{
		var diff = measure1 - measure2;
		return diff > 5 || diff < -5;
	}
}

public class VervNavnComparer : IEqualityComparer<(string, string)>
{
	public bool Equals((string, string) x, (string, string) y)
	{
		return x.Item1.Equals(y.Item1);
	}

	public int GetHashCode((string, string) obj)
	{
		return obj.Item1.GetHashCode();
	}
}

public class Medlem
{
	public string Navn { get; set; }
	public string Tlf { get; set; }
	public string Epost { get; set; }
}

public class MedlemsRegister
{
	private string swFolder;
	private string swEksportFil;

	public Dictionary<string, Medlem> Medlemmer { get; }
	
	public MedlemsRegister()
	{
		swFolder = @"C:\MyLocal\Solviken\FraStyreWeb";
		swEksportFil = Path.Combine(swFolder, "Standard_Rapport.csv");
		Medlemmer = new Dictionary<string, Medlem>();
	}
	
	public MedlemsRegister LesData()
	{
		var downloadFolder = @"C:\Users\solvi\Downloads";
		CopyNewerFile(Path.Combine(downloadFolder, "Standard_Rapport.csv"), swFolder);
		using (var reader = new StreamReader(swEksportFil, Encoding.GetEncoding("UTF-8")))
		{
			reader.ReadLine();      // Skip header
			string line;
			while ((line = reader.ReadLine()) != null)
			{
				var fields = line.Split('\t');
				if (fields.Length >= 5)
				{
					var name = $"{fields[1]} {fields[0]}";
					var tlf = fields[4];
					Medlemmer[name] = new Medlem {Navn = name, Tlf = tlf};
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