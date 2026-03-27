<Query Kind="Program">
  <Reference>&lt;RuntimeDirectory&gt;\System.Collections.Concurrent.dll</Reference>
  <NuGetReference>ExcelDataReader</NuGetReference>
  <NuGetReference>ExcelDataReader.DataSet</NuGetReference>
  <Namespace>ExcelDataReader</Namespace>
  <Namespace>System</Namespace>
  <Namespace>System.Collections.Concurrent</Namespace>
  <Namespace>System.Data</Namespace>
  <Namespace>System.Globalization</Namespace>
  <Namespace>System.IO</Namespace>
  <Namespace>System.Net</Namespace>
</Query>

void Main(string[] args)
{
	bool bryggeliste = false;
	string path = null;
	string prefix = null;
	string fromDate = null;
	
	if (args?.Length > 0)
	{
		// [-file outfile] [-bryggeliste] [-prefix prefix] [-fromdate fromdate]
		for (int i = 0; i < args.Length; i++)
		{
			var arg = args[i];
			if (arg.Equals("-bryggeliste", StringComparison.OrdinalIgnoreCase))
			{
				bryggeliste = true;
			}
			else if (arg.Equals("-file", StringComparison.OrdinalIgnoreCase))
			{
				path = args[++i];
			}
			else if (arg.Equals("-prefix", StringComparison.OrdinalIgnoreCase))
			{
				prefix = args[++i];
			}
			else if (arg.Equals("-fromdate", StringComparison.OrdinalIgnoreCase))
			{
				fromDate = args[++i];
			}
		}
	}

//
//
	//Enumerable.Range(1, 6)
	//	.ToList()
	//	.ForEach(e => VisAlledata(new StyreWebExport().LesData(e.ToString()), true, @"C:\MyLocal\Solviken\Rapporter"));

	VisAlleData(new StyreWebExport().LesData(prefix, fromDate), bryggeliste);
	if (path != null)
	{
		VisAlleData(new StyreWebExport().LesData(prefix, fromDate), bryggeliste, path);
	}

	//VisAlledata(new ExcelExport().LesData(), bryggeliste);
	//VisAlledata(new HavneWebExport().LesData(), bryggeliste);

	//VisEierEndringer(new StyreWebExport().LesData(fromDate: "21.08.2025"),
	//				 new StyreWebExport().LesData());
	//VisEierEndringer(new HavneWebExport().LesData(), new StyreWebExport().LesData());
	//VisSluttedeEiere(new HavneWebExport().LesData(), new StyreWebExport().LesData());
	//VisEierEndringer(new ExcelExport().LesData(), new StyreWebExport().LesData());
	//new List<int>{1, 2, 3, 5, 6}.ForEach(x => VisArealForskjeller(new HavneWebExport().LesData(x.ToString()), new StyreWebExport().LesData(x.ToString())));
	//VisArealForskjeller(new HavneWebExport().LesData("6"), new StyreWebExport().LesData("6"));
	//VisVaktFritak(new HavneWebExport().LesData());
	//BeregnBatplassAvgifter(new StyreWebExport().LesData());
	//VisAlleMedVaktplikt(new StyreWebExport().LesData(), new HavneWebExport().LesData());
	//VisLedigePlasser(new StyreWebExport().LesData());
	//VisAlleMedVaktfritakOgPlasser(new StyreWebExport().LesData(fromDate: "31.08.2025"));
	//VisAlleMedVaktfritakOgPlasser(new StyreWebExport().LesData());
	//SjekkSesongOgUngdom(new StyreWebExport().LesData());
	//FinnLeietillegg(new StyreWebExport().LesData());
	//SjekkVareVarianter(new StyreWebExport().LesData());
	//FinnPlasserUnder2500(new StyreWebExport().LesData());
	//FinnEierEndringerEtter(DateTime.Parse("20.06.2025"), new StyreWebExport().LesData());
	//FinnSesongLeiereFraAndelsplass(new StyreWebExport().LesData());
	//SammenlignGrupper(new HavneWebExport().LesData(), new StyreWebExport().LesData());
	//SammenlignBatplassAvgift(new StyreWebExport().LesData(fromDate: "21.01.2026"),  // Siste dato før omlegging av bredde/lengde
	//						new StyreWebExport().LesData());
	//BeregnBesteGrenserogVerdier(new StyreWebExport().LesData(fromDate: "21.01.2026"),  // Siste dato før omlegging av bredde/lengde
	//						new StyreWebExport().LesData());
	//UploadFtp("ftp://vd03.verdidata.no/StyreWebStatus.html", @"C:\Users\Solviken\Downloads\StyreWebStatus.html", "solvikenftp", "c5712$vQp");
}

void SammenlignBatplassAvgift(HavneData havn1, HavneData havn2)
{
	int gammelTotal = 0;
	int nyTotal = 0;
	int[] gammelGruppeTotal = new int[7];
	int[] nyGruppeTotal = new int[7];
	int[] antall = new int[7];
	int prismodell = 3;

	var writer = new StreamWriter(@"C:\Users\Solviken\Downloads\NyeAvgifter.csv", false, Encoding.GetEncoding("UTF-8"));
	Console.SetOut(writer);

	Console.WriteLine("Forslag til nye avgiftsgrupper for båtplasser:");
	for (int i = 0; i < Priser.Prisgrupper[prismodell].Length; i++)
	{
		Console.WriteLine($"Gruppe {i + 1}:\t< {Priser.Prisgrupper[prismodell][i].limit}m\t{Priser.Prisgrupper[prismodell][i].factor} kr/m");
	}
	Console.WriteLine();
	Console.WriteLine("Plass\tBredde m\tPris/m\tFørpris\tNy pris\tEndring\tProsent");
	
	foreach (var plass2 in havn2.GetAlleBryggePlasser()
			.Where(p => !p.UngdomsPlass && !p.JollePlass && p.PlassId.Length == 4))
	{
		var plass1 = havn1.GetBatPlass(plass2.PlassId);
		if (plass1 != null)
		{
			int gammelAvgift = plass1.BeregnBatplassAvgift();
			int nyavgift = plass2.BeregnNyBatplassAvgift(prismodell);
			gammelTotal += gammelAvgift;
			nyTotal += nyavgift;
			int gruppe = plass2.GetPrisGruppe();
			antall[gruppe]++;
			gammelGruppeTotal[gruppe] += gammelAvgift;
			nyGruppeTotal[gruppe] += nyavgift;
			int diff = nyavgift - gammelAvgift;
			var prosent = Math.Round(((double)diff / gammelAvgift) * 100);
			var bredde = plass2.Bredde / 100.0;
			int faktor = plass2.GetPrisFaktor(prismodell);
			Console.WriteLine($"{plass1.PlassId}:\t{bredde}\t{faktor}\t{gammelAvgift}\t{nyavgift}\t{diff}\t{prosent}");
		}
	}

	Console.WriteLine($"\nGammel total:\t{gammelTotal}");
	Console.WriteLine($"Ny total:\t{nyTotal}");
	Console.WriteLine($"Endring:\t{nyTotal - gammelTotal}");
	Console.WriteLine("\nPr. gruppe:");
	Console.WriteLine("Gruppe\tAntall\tTotal før\tTotal nå\tEndring");
	
	for (int i = 0; i < 7; i++)
	{
		Console.WriteLine($"Gruppe {i + 1}:\t{antall[i]}\t{gammelGruppeTotal[i]}\t{nyGruppeTotal[i]}\t{(nyGruppeTotal[i] - gammelGruppeTotal[i])}");
	}
	
	writer.Close();
}

void BeregnBesteGrenserogVerdier(HavneData havn1, HavneData havn2)
{
	var gamlePriser = new Dictionary<string, int>();
	(double limit, int factor)[] prisgrupper =
		new[] { (2.5, 800), (3.0, 1005), (3.5, 1050), (4.0, 1420), (4.5, 2000), (5.0, 2200), (10.0, 2500) };
	int[] besteFaktor = new int[7];
	var grupper = new List<BatPlass>[7];
	int[] gammelTotal = new int[7];
	
	for (int i = 0; i < grupper.Length; i++)
	{
		grupper[i] = new List<BatPlass>();
	}

	foreach (var plass in havn2.GetAlleBryggePlasser()
		.Where(p => (p.AndelsPlass || p.SesongPlass || p.FramleiePlass) && p.PlassId.Length == 4))
	{
		var gammelPlass = havn1.GetBatPlass(plass.PlassId);
		if (gammelPlass != null)
		{
			var gammelPris = gamlePriser[gammelPlass.PlassId] = gammelPlass.BeregnBatplassAvgift();
			double bredde = (double)plass.Bredde / 100.0;

			for (int i = 0; i < prisgrupper.Length; i++)
			{
				if (bredde < prisgrupper[i].limit)
				{
					grupper[i].Add(plass);
					gammelTotal[i] += gammelPris;
					break;
				}
			}
		}
	}

	for (int prisgruppe = 0; prisgruppe < prisgrupper.Length; prisgruppe++)
	{
		int suggestedFactor = prisgrupper[prisgruppe].factor;
		//int minsteProsent = 10000;
		for (int factor = suggestedFactor / 2; factor < suggestedFactor * 2; factor += 5)
		{
			//int totalProsent = 0;
			//int totalDiff = 0;
			int antallPlasser = grupper[prisgruppe].Count;
			if (antallPlasser > 0)
			{
				int totPris = 0;
				foreach (var plass in grupper[prisgruppe])
				{
					double bredde = (double)plass.Bredde / 100.0;
					int pris = (int)(bredde * factor);
					totPris += pris;
					//int gammelPris = gamlePriser[plass.PlassId];
					//int diff = pris - gammelPris;
					//totalDiff += diff;
					//int prosent = (int)Math.Round(((double)diff / gammelPris) * 100.0);
					//totalProsent += prosent;
				}

				if (totPris >= gammelTotal[prisgruppe])
				{
					// Denne faktoren er "god nok" for prisgruppen
					besteFaktor[prisgruppe] = factor;
					break;
				}
				//if (totalProsent >= 0)
				//{
				//	// Se bort fra faktorer som gir lavere båtplassavgift enn før
				//	int snittProsent = totalProsent / antallPlasser;
				//	if (snittProsent < minsteProsent)
				//	{
				//		minsteProsent = snittProsent;
				//		besteFaktor[prisgruppe] = factor;
				//	}
				//}
			}
		}
	}

	Console.WriteLine("Beste faktorer:");
	for (int i = 0; i < prisgrupper.Length; i++)
	{
		Console.WriteLine($"< {prisgrupper[i].limit}: {besteFaktor[i]}");
	}
}

void VisSluttedeEiere(HavneData havn1, HavneData havn2)
{
	var eiere2 = new HashSet<string>();
	foreach (var plass in havn2.GetAndelsPlasser())
	{
		if (!string.IsNullOrEmpty(plass.Eier))
		{
			eiere2.Add(plass.Eier);
		}
	}
	
	foreach (var plass in havn1.GetAllePlasser())
	{
		if (!string.IsNullOrEmpty(plass.Eier))
		{
			if (!eiere2.Contains(plass.Eier))
			{
				Console.WriteLine($"{plass.PlassId}: {plass.Eier} har sluttet");
			}
		}
	}
}

void SammenlignGrupper(HavneData havn1, HavneData havn2)
{
	foreach (var plass1 in havn1.GetAllePlasser().Where(p => !p.LandOpplag))
	{
		var plass2 = havn2.GetBatPlass(plass1.PlassId);
		if (plass2 == null)
		{
			Console.WriteLine($"{plass1.PlassId} er ikke i {havn2.Navn}");
			continue;
		}
		
		if (plass1.Gruppe != plass2.Gruppe)
		{
			Console.WriteLine($"Forskjellig gruppe for {plass1.PlassId}: {havn1.Navn}: '{plass1.Gruppe}', {havn2.Navn}: '{plass2.Gruppe}'");
		}
	}
}

void FinnSesongLeiereFraAndelsplass(HavneData havn)
{
	var andelsplasser = havn.GetAndelsPlasser();
	var leietakere = new List<(string, string, string, string)>();		// Navn, etternavn, plass, eier
	foreach (var andelsplass in andelsplasser)
	{
		if (andelsplass.Leier != null)
		{
			var etternavn = andelsplass.Leier.Split(' ').Reverse().ElementAt(0);
			leietakere.Add((andelsplass.Leier, etternavn, andelsplass.PlassId, andelsplass.Eier));
		}
	}
	
	leietakere.Sort((p1, p2) => p1.Item2.CompareTo(p2.Item2));
	
	Console.WriteLine("Leietakere til andelsplass:");
	Console.WriteLine("Leietaker              Plass     Eier");
	foreach (var leietaker in leietakere)
	{
		Console.WriteLine($"{leietaker.Item1.PadRight(22)} {leietaker.Item3}      {leietaker.Item4}");
	}
}

void FinnEierEndringerEtter(DateTime time, HavneData havn)
{
	// Antar at alt fram til "time" er fakturert. Finn endringer siden det som skal faktureres
	// Er plasser tildelt etter time?
	var nyeTildelinger = havn.GetAndelsPlasser().Where(h => h.UtlevertFra > time).ToList();
	Console.WriteLine($"Andelsplasser tildelt etter {time.ToShortDateString()}:");
	Console.WriteLine();
	
	foreach (var plass in nyeTildelinger)
	{
		Console.WriteLine($"{plass.PlassId}: {plass.UtlevertFra.ToShortDateString()} - {plass.Eier}");
	}

	var nyeSesongplasser = havn.GetSesongPlasser().Where(h => h.UtLeidFra > time);
	Console.WriteLine();
	Console.WriteLine($"Sesongplasser tildelt etter {time.ToShortDateString()}:");
	Console.WriteLine();
	foreach (var plass in nyeSesongplasser)
	{
		Console.Write($"{plass.PlassId}: {plass.UtLeidFra.ToShortDateString()} - {plass.Leier.PadRight(20)}");
		//Console.WriteLine($"- fra {plass.Eier}");

		if (plass.Eier != null)
		{
			Console.WriteLine($"- fra {plass.Eier}");
		}
		else
		{
			Console.WriteLine();
		}
	}

	var nyeUngdomsplasser = havn.GetUngdomsPlasser().Where(h => h.UtLeidFra > time);
	Console.WriteLine();
	Console.WriteLine($"Ungdomsplasser tildelt etter {time.ToShortDateString()}:");
	Console.WriteLine();
	foreach (var plass in nyeUngdomsplasser)
	{
		Console.Write($"{plass.PlassId}: {plass.UtLeidFra.ToShortDateString()} - {plass.Leier.PadRight(20)}");
		//Console.WriteLine($"- fra {plass.Eier}");

		if (plass.Eier != null)
		{
			Console.WriteLine($"- fra {plass.Eier}");
		}
		else
		{
			Console.WriteLine();
		}
	}
}

void FinnPlasserUnder2500(HavneData havn)
{
	var plasser = havn.GetAndelsPlasser().Concat(havn.GetSesongPlasser());
	var smaPlasser = new List<BatPlass>();
	foreach (var plass in plasser)
	{
		var lengde = ((double)plass.Lengde) / 100;
		var bredde = ((double)plass.Bredde) / 100;
		
		int beregnetAvgift = (int)Math.Round(bredde * lengde * plass.PrisFaktor(lengde));
		if (beregnetAvgift < 2500)
		{
			smaPlasser.Add(plass);
		}
	}

	var eiere = smaPlasser.Select(p => p.Leier ?? p.Eier).OrderBy(e => e).ToList();
	Console.WriteLine($"Båteiere med plasser under minimum avgift ({smaPlasser.Count} stk.)\n");
	foreach (var eier in eiere)
	{
		Console.WriteLine(eier);
	}
}

void SjekkVareVarianter(HavneData havn)
{
	var plasser = havn.GetAndelsPlasser().Concat(havn.GetSesongPlasser());
	foreach (var plass in plasser)
	{
		var bredde = ((double)plass.Bredde) / 100;
		var beregnetVareVariant = VareVariant.Create(bredde);
		if (plass.VareVariant.Gruppe != beregnetVareVariant.Gruppe)
		{
			Console.WriteLine($"{plass.PlassId}: Feil varevariant {plass.VareVariant.Text}, skal være {beregnetVareVariant.Text}");
		}
	}
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

void VisAlleMedVaktfritakOgPlasser(HavneData havn, string file = null)
{
	var fritaksPlasser = havn.GetAllePlasser().Where(p => p.Vaktfritak != null);
	var sortert = fritaksPlasser.OrderBy(p => p.Bruker);
	var gruppert = sortert.GroupBy(s => s.Bruker);
	var sortertPaArsak = new Dictionary<string, List<IGrouping<string, BatPlass>>>();		// Årsak, brukere
	Console.Write("\n+Medlemmer med fritak for vakt og dugnad");
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

		Console.WriteLine($"Antall fritaksplasser: {arsakPlasser}");
	}
	
	Console.WriteLine($"\nAntall medlemmer med vakt/dugnadsfritak: {gruppert.Count()}, båtplasser: {antallFritak}");
	Console.WriteLine("-");
	
	// Skal finne alle medlemmer med dugnadsplikt, og hvor mange dugnadstimer (båtplasser x 8)
	var dugnadsplikt = new ConcurrentDictionary<string, (int, string)>();		// navn, timer
	var andelsplasser = havn.GetAndelsPlasser();
	var sesongplasser = havn.GetSesongPlasser();
	var pliktigeplasser = andelsplasser.Concat(sesongplasser).Distinct().OrderBy(a => a.PlassId);

	// Havneinstruks punkt 12: Ved fylte 75 år halveres dugnadsplikt
	// For 2026 vil det si født i 1951 eller tidligere
	var alder75pluss = new HashSet<string>();
	var fil75pluss = @"C:\MyLocal\Solviken\FraStyreWeb\Dato_Alder_start_slutt.csv";		// 
	using (var reader = new StreamReader(fil75pluss))
	{
		reader.ReadLine();
		string line;
		while ((line = reader.ReadLine()) != null)
		{
			var parts = line.Split('\t');
			if (parts.Length < 3)
			{
				break;
			}

			var navn = $"{parts[2]} {parts[1]}";
			alder75pluss.Add(navn);
		}
	}

	var sisteDugnadSpring2026 = DateTime.Parse("23.05.2026");
	foreach (var plass in pliktigeplasser)
	{
		if (plass.Vaktfritak == null)
		{
			var key = plass.Leier ?? plass.Eier;
			var utlevert = plass.Leier != null ? plass.UtLeidFra : plass.UtlevertFra;
			int dugnadsPliktTimer = utlevert > sisteDugnadSpring2026 ? 4 : 8;
			if (alder75pluss.Contains(key))
			{
				dugnadsPliktTimer /= 2;
			}
			
			dugnadsplikt.AddOrUpdate(
				key, 
				(dugnadsPliktTimer, plass.PlassId), 
				(k, v) => (v.Item1 + dugnadsPliktTimer, v.Item2 + "," + plass.PlassId)
			);
		}
	}
	
	var dugnadsPliktListe = dugnadsplikt.OrderBy(d => d.Key.Split(' ').Reverse().ToArray()[0]).ToList();	// Sorter på etternavn
	int totalTimer = 0;
	int totalVakter = 0;
	Console.WriteLine("\n+Vakt/dugnadspliktige");
	Console.WriteLine("Medlem                      Vakter  Dugnad   Plass(er)");
	foreach (var pliktig in dugnadsPliktListe)
	{
		string medlem = pliktig.Key;
		int timer = pliktig.Value.Item1;
		totalTimer += timer;
		string plasser = pliktig.Value.Item2;
		int vakter = plasser.Count(p => p == ',') + 1;
		totalVakter += vakter;
		Console.WriteLine($"{pliktig.Key, -30}{vakter, 2}{timer,6}       {plasser, -20}");
	}

	Console.WriteLine($"\nTotalt {totalVakter} pliktige vakter");
	Console.WriteLine($"Totalt {totalTimer} pliktige dugnadstimer");
	Console.WriteLine("-");

	// Les export av dugnadsregnskap 2026
	var dugnadsRegnskap = new ConcurrentDictionary<string, int>();
	var dugnadFileName = @"C:\MyLocal\Solviken\FraStyreWeb\Dugnadsregnskap.csv";
	int totalInnsats = 0;
	
	using (var reader = new StreamReader(dugnadFileName))
	{
		reader.ReadLine();
		reader.ReadLine();
		reader.ReadLine();
		string line;
		while ((line = reader.ReadLine()) != null)
		{
			var fields = line.Split('\t');
			if (string.IsNullOrEmpty(fields[0]))
			{
				break;
			}

			if (fields.Length == 5 && int.TryParse(fields[4], out var innsats))
			{
				var navn = $"{fields[1]} {fields[0]}";
				dugnadsRegnskap.AddOrUpdate(navn, innsats, (n, v) => v + innsats);
				totalInnsats += innsats;
			}
		}
	}

	var innsatsListe = dugnadsRegnskap.OrderBy(d => d.Key.Split(' ').Reverse().ToArray()[0]).ToList();
	Console.WriteLine("\n+Registrerte dugnadstimer i 2026");
	innsatsListe.ForEach(l => Console.WriteLine($"{l.Key,-30}{l.Value, 2}"));
	
	var manglendeTimer = new List<(string, int)>();
	
	foreach (var dugnadsPlikt in dugnadsPliktListe)
	{
		int pliktigeTimer = dugnadsPlikt.Value.Item1;
		if (dugnadsRegnskap.TryGetValue(dugnadsPlikt.Key, out var timer))
		{
			pliktigeTimer -= timer;
			innsatsListe.RemoveAt(innsatsListe.IndexOf(new KeyValuePair<string, int>(dugnadsPlikt.Key, timer)));
		}
		
		if (pliktigeTimer > 0)
		{
			manglendeTimer.Add((dugnadsPlikt.Key, pliktigeTimer));
		}
	}
	
	// Utført dugnad som ikke er satt på båtplass
	var ukjente = new List<string>();
	foreach (var utfort in innsatsListe)
	{
		if (!fritaksPlasser.Any(p => (p.Leier ?? p.Eier) == utfort.Key))
		{
			string utfortFor = FinnUtfortFor(utfort.Key);
			if (utfortFor != null)
			{
				int ix = manglendeTimer.FindIndex(a => a.Item1 == utfortFor);
				if (ix >= 0)
				{
					var mangler = manglendeTimer[ix];
					manglendeTimer[ix] = (mangler.Item1, mangler.Item2 - utfort.Value);
					ukjente.Add($"{utfort.Key, -30} {utfort.Value, 10} (for {mangler.Item1})");
				}
				else
				{
					ukjente.Add($"Fant ikke {utfortFor} i listen over manglende dugnadstimer");
				}
			}
			else
			{
				ukjente.Add($"{utfort.Key,-30} {utfort.Value, 10}");
			}
		}
	}

	if (ukjente.Count > 0)
	{
		Console.WriteLine("\nUtført dugnad uten plikt, skal føres på andre?:");
		foreach (var ukjent in ukjente)
		{
			Console.WriteLine(ukjent);
		}
	}
	Console.WriteLine($"\nAntall timer utført dugnad i 2026: {totalInnsats}");
	Console.WriteLine("-");

	Console.WriteLine("\n+Manglende dugnadstimer i 2026");
	int totalMangler = 0;
	//var gruppeFilNavn = @"C:\MyLocal\Solviken\FraStyreWeb\GruppeImportDugnadsFaktura.csv";
	//using (var writer = new StreamWriter(gruppeFilNavn))
	{
		//writer.WriteLine("Visningsnavn");
		foreach (var mangler in manglendeTimer)
		{
			if (mangler.Item2 > 0)
			{
				totalMangler += mangler.Item2;
				Console.WriteLine($"{mangler.Item1,-30}{mangler.Item2,2}");
				//writer.WriteLine($"{mangler.Item1}\t{mangler.Item2}");
			}
		}
	}

	Console.WriteLine($"\nAntall timer ikke utført dugnad i 2026: {totalMangler}");
	Console.WriteLine("-");
}

string FinnUtfortFor(string key)
{
	switch (key)
	{
		case "Morten Dundas":
			return "Carl Edvard Reinertsen";
		case "Frode Spangelo":
			return "Øystein Spangelo";
		// Fyll på med flere ...
	}
	
	return null;
}

void VisLedigePlasser(HavneData havn)
{
	var ledige = havn.GetLedigePlasser();
	var sortert = ledige.OrderBy(l => l.Bredde);
	foreach (var plass in sortert)
	{
		Console.WriteLine($"{plass.PlassId}: {plass.Bredde} m");
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
	foreach (var plass in havneData.GetAndelsPlasser()
				.Concat(havneData.GetSesongPlasser()
				.Concat(havneData.GetUngdomsPlasser())))
	{
		var batplassAvgift = plass.BeregnBatplassAvgift();
		totalFakturering += batplassAvgift;
		Console.WriteLine($"{plass.PlassId}: {(plass.Leier ?? plass.Eier),-30} (BxL: {plass.Bredde}x{plass.Lengde}) kr. {batplassAvgift:n}");
	}

	Console.WriteLine($"\nTotal fakturering av båtplassavgifter i 2025: kr. {totalFakturering:n}");
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
				if (plass1.Bredde.Ulik(plass2.Bredde) || plass1.Lengde.Ulik(plass2.Lengde))
				{
					var bredde1 = ((double)plass1.Bredde / 100).ToString("0.00");
					var lengde1 = ((double)plass1.Lengde / 100).ToString("0.00");
					var bredde2 = ((double)plass2.Bredde / 100).ToString("0.00");
					var lengde2 = ((double)plass2.Lengde / 100).ToString("0.00");
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
	var tapere = new List<(string, string, bool)>();      // Eier, plassId, is_matched_with_winner
	var vinnere = new List<(string, string, bool)>();      // Eier, plassId, is_matched_with_loser

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
				tapere.Add((taper, plass1.PlassId, false));
				vinnere.Add((vinner, plass2.PlassId, false));
				
				Console.WriteLine($"{plass1.PlassId}: Eier fra {taper} til {vinner}");
			}
		}
		else
		{
			Console.WriteLine($"{plass1.PlassId} mangler i {havn2.Navn}");
		}
	}
	
	Console.WriteLine("\nInnskudd som skal krediteres (Merket * har hatt flere plasser, må sjekkes:");
	var venteListe = new List<string>();
	var flytteListe = new List<string>();
	foreach (var taper in tapere)
	{
		int i;
		string merket = string.Empty;
		for (i = 0; i < vinnere.Count; i++)
		{
			var vinner = vinnere[i];
			if (vinner.Item1 == taper.Item1)
			{
				if (vinner.Item3)
				{
					merket = "*";
				}
				else
				{
					vinnere[i] = (vinner.Item1, vinner.Item2, true);
					break;
				}
			}
		}

		var gammelPlass = havn1.GetBatPlass(taper.Item2);
		var gammeltInnskudd = gammelPlass.Innskudd.ToString("n").PadLeft(9);
		var nyPlass = havn2.GetBatPlass(taper.Item2);
		
		if (i == vinnere.Count)
		{
			// Gitt fra seg plass uten å få ny
			if (nyPlass.Eier != null)
			{
				Console.WriteLine($"{taper.Item2}: kr. {gammeltInnskudd} - {taper.Item1}{merket}");
			}
			else
			{
				venteListe.Add($"{taper.Item2}: kr. {gammeltInnskudd} - {taper.Item1}{merket}");
			}
		}
		else if (taper.Item1 != "Ledig")
		{
			nyPlass = havn2.GetBatPlass(vinnere[i].Item2);
			// Byttet plass i havna
			// Sjekk om innskudd skal økes
			var nyttInskudd = havn2.BeregnInnskudd(nyPlass.PlassId);
			var innskuddDiff = string.Empty;
			if (gammelPlass.Innskudd < nyttInskudd.Item1)
			{
				innskuddDiff = $" - Innskudd økes med {nyttInskudd.Item1 - gammelPlass.Innskudd}";
			}
			flytteListe.Add($"{taper.Item1.PadRight(22)} fra {taper.Item2} til {nyPlass.PlassId}{innskuddDiff}");
		}
	}
	
	Console.WriteLine("\nInnskudd som krediteres etter at plassen er videresolgt:");
	foreach (var vente in venteListe)
	{
		Console.WriteLine(vente);
	}

	Console.WriteLine("\nNye innskudd (Merket * har hatt flere plasser, må sjekkes:");
	foreach (var vinner in vinnere)
	{
		int i;
		string merket = string.Empty;
		for (i = 0; i < tapere.Count; i++)
		{
			var taper = tapere[i];
			if (taper.Item1 == vinner.Item1)
			{
				if (taper.Item3)
				{
					merket = "*";
				}
				else
				{
					tapere[i] = (taper.Item1, taper.Item2, true);
					break;
				}
			}
		}

		if (i == tapere.Count && vinner.Item1 != "Ledig")
		{
			// Fått ny plass uten å gi fra seg en
			var innskudd = havn2.BeregnInnskudd(vinner.Item2);
			Console.Write($"{vinner.Item2}: kr. {innskudd.Item1.ToString("n").PadLeft(9)} - {vinner.Item1}{merket}");
			if (havn2.GetBatPlass(vinner.Item2).Innskudd == innskudd.Item1)
			{
				Console.WriteLine(" (Betalt)");
			}
			else
			{
				Console.WriteLine();
			}
		}

	}

	Console.WriteLine("\nBåtplassbytter:");
	foreach (var vente in flytteListe)
	{
		Console.WriteLine(vente);
	}
}

void VisAlleData(HavneData dataSet, bool bryggeliste = true, string file = null)
{
	var andelsplasser = dataSet.GetAndelsPlasser();
	var sesongplasser = dataSet.GetSesongPlasser();
	var framleiePlasser = dataSet.GetFramleiePlasser();
	var ungdomsplasser = dataSet.GetUngdomsPlasser();
	var jollePlasser = dataSet.GetJollePlasser();
	var tilLeiePlasser = dataSet.GetTilLeiePlasser();
	var ledigePlasser = dataSet.GetLedigePlasser();
	var batplassVenteListe = dataSet.BatplassVenteliste;
	var innskuddVenteliste = dataSet.InnskuddVenteliste;
	var medlemsRegister = new MedlemsRegister().LesData();
	TextWriter originalOut = null;
	StreamWriter writer = null;
	
	if (file != null)
	{
		// Save the original output stream
		originalOut = Console.Out;
		writer = new StreamWriter(file, false, Encoding.GetEncoding("UTF-8"));
		Console.SetOut(writer);
	}

	Console.Write($"Eksport fra {dataSet.Navn} {dataSet.TimeStamp}");
	if (dataSet.PlassPrefix != null)
	{
		Console.WriteLine($" - Plasser som starter med \"{dataSet.PlassPrefix}\"");
	}
	else
	{
		Console.WriteLine();
	}

	var feil = SjekkForFeil(dataSet, innskuddVenteliste, ledigePlasser);
	if (feil.Count > 0)
	{
		Console.WriteLine("\n+ Feil i båtplassdata");
		feil.ForEach(f => Console.WriteLine(f));
	}
	Console.WriteLine("-");
	
	if (bryggeliste)
	{
		var allePlasser = dataSet.GetAllePlasser();
		Console.WriteLine($"\n{allePlasser.Count} båtplasser");
		foreach (var plass in allePlasser)
		{
			PrintBatplass2(plass, medlemsRegister);
		}

		if (writer != null)
		{
			writer.Close();
		}
		
		return;
	}

	Console.WriteLine("\n+Andelsplasser");
	Console.WriteLine($"{andelsplasser.Count} andelsplasser");
	foreach (var plass in andelsplasser)
	{
		PrintBatplass(plass, true, medlemsRegister);
	}
	Console.WriteLine("-");

	Console.WriteLine("\n+Framleide plasser");
	Console.WriteLine($"{framleiePlasser.Count} framleide plasser");
	foreach (var plass in framleiePlasser)
	{
		var eier = $"(fra {plass.Eier})";
		PrintBatplass(plass, false, medlemsRegister, eier);
	}
	Console.WriteLine("-");

	Console.WriteLine("\n+Sesongplasser");
	Console.WriteLine($"{sesongplasser.Count} sesongplasser");
	foreach (var plass in sesongplasser)
	{
		if (plass.Leier != null)
		{
			// 2025-style marina, sesongplass er framleie
			var eier = plass.Eier != null ? $"(fra {plass.Eier})" : null;
			PrintBatplass(plass, false, medlemsRegister, eier);
		}
		else
		{
			// 2026-style, sesongplass ikke framleie
			PrintBatplass(plass, true, medlemsRegister, null);
		}
	}
	Console.WriteLine("-");

	if (ungdomsplasser.Count > 0)
	{
		Console.WriteLine("\n+Ungdomsplasser");
		Console.WriteLine($"{ungdomsplasser.Count} ungdomsplasser");
		foreach (var plass in ungdomsplasser)
		{
			Console.WriteLine($"{plass.PlassId}: {plass.Bruker,-25}");
		}
		Console.WriteLine("-");
	}

	Console.WriteLine("\n+Jolleplasser");
	Console.WriteLine($"{jollePlasser.Count} jolleplasser");
	foreach (var plass in jollePlasser)
	{
		Console.WriteLine($"{plass.PlassId}: {plass.Bruker,-25}");
	}
	Console.WriteLine("-");

	// Til leie-plasser
	Console.WriteLine("\n+Plasser til leie");
	Console.WriteLine($"{tilLeiePlasser.Count} plasser til leie");
	foreach (var plass in tilLeiePlasser)
	{
		PrintBatplass(plass, true, medlemsRegister);
	}
	Console.WriteLine("-");

	Console.WriteLine("\n+Ledige plasser");
	Console.WriteLine($"{ledigePlasser.Count} ledige plasser");
	var breddeListe = new List<(string, string, char)>();
	var ledigePrGruppe = new Dictionary<char, int>
	{
		{'A', 0},
		{'B', 0},
		{'C', 0},
		{'D', 0},
		{'E', 0},
		{'F', 0},
		{'G', 0},
		{'H', 0},
		{'L', 0},
	};

	foreach (var plass in ledigePlasser)
	{
		var breddeMeter = (plass.Bredde / 100.0).ToString();
		Console.WriteLine($"{plass.PlassId}: {breddeMeter,10} m     Gruppe {plass.Gruppe}");
		breddeListe.Add((plass.PlassId, breddeMeter, plass.Gruppe));
		ledigePrGruppe[plass.Gruppe]++;
	}
	
	var sortert = breddeListe.OrderBy(l => l.Item2);
	
	Console.WriteLine("\nLedige plasser sortert på bredde:");
	foreach (var plass in sortert)
	{
		Console.WriteLine($"{plass.Item1}: {plass.Item2, 10} m     Gruppe {plass.Item3}");
	}

	Console.WriteLine("\nLedige plasser pr. gruppe:");
	foreach (var ledige in ledigePrGruppe)
	{
		Console.WriteLine($"{ledige.Key}: {ledige.Value}");
	}
	Console.WriteLine("-");

	Console.WriteLine("\n+Innskudd ikke tilbakebetalt");
	//Console.WriteLine($"\n*** Innskudd som ikke er tilbakebetalt ({innskuddVenteliste.Count()}) ***\n");
	Console.WriteLine("Medlem                    Plass    Innskudd     Solgt   Ønsker tilbakebetaling");
	Console.WriteLine("--------------------------------------------------------------------------");
	
	int total = 0;
	foreach (var innskudd in innskuddVenteliste)
	{
		var navn = innskudd.Navn;
		var plassId = innskudd.PlassId;
		var kroner = innskudd.Innskudd;
		total += kroner;
		var solgt = dataSet.GetBatPlass(innskudd.PlassId).AndelsPlass;
		var utbetales = innskudd.Utbetales;
		
		Console.WriteLine($"{navn,-25} {plassId} {kroner,10}        {(solgt ? "J" : "N")}       {(innskudd.Utbetales ? "J" : "N")}");
	}

	Console.WriteLine($"\nInnskudd som venter på utbetaling: {total} kr");
	Console.WriteLine("-");

	PrintVentelister(batplassVenteListe, andelsplasser);
	PrintLandopplag(dataSet);
	VisAlleMedVaktfritakOgPlasser(dataSet);
	
	string venteListeFil = null;
	if (file != null)
	{
		// Denne skal ut på separat fil i tillegg
		venteListeFil = Path.Combine(Path.GetDirectoryName(file), "venteliste.txt");
		using (var ventelisteWriter = new StreamWriter(venteListeFil, false, Encoding.GetEncoding("UTF-8")))
		{
			Console.SetOut(ventelisteWriter);
			Console.WriteLine($"Solviken Båtforening, status fra StyreWeb {DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss")}");
			PrintVentelister(batplassVenteListe, andelsplasser);
		}
		writer.Close();

		Console.SetOut(originalOut);
		PublishStatus(file, true);
		PublishStatus(venteListeFil);
	}
}

void PrintLandopplag(HavneData dataSet)
{
	Console.WriteLine("\n+Landopplagsplasser");
	dataSet.GetLandOpplagsPlasser().ForEach(s => Console.WriteLine($"{s.PlassId}: {s.Eier??"Ledig",-30}"));
	Console.WriteLine("-");
}

List<string> SjekkForFeil(HavneData dataSet, List<InnskuddEier> innskuddVenteliste, List<BatPlass> ledigePlasser)
{
	var result = new List<string>();
	foreach (var plass in dataSet.GetAlleBryggePlasser())
	{
		if (plass.Eier == null && (plass.AndelsPlass || plass.SesongPlass || plass.JollePlass))
		{
			result.Add($"{plass.PlassId}: Ingen eier, men plassen er markert i bruk");
		}

		if (plass.AndelsPlass && plass.Innskudd == 0)
		{
			result.Add($"{plass.PlassId}: Andelsplass uten innskudd");
		}
		else if (plass.Innskudd != 0 && !plass.AndelsPlass)
		{
			result.Add($"{plass.PlassId}: Innskudd på plass som ikke er andelsplass");
		}

		// Sjekk framleieplasser
		if (plass.FramleiePlass && (plass.Eier == null || plass.Leier == null))
		{
			result.Add($"{plass.PlassId}: Framleieplass med feil i eier eller leietaker");
		}
		else if (plass.Eier != null && plass.Leier != null && !plass.FramleiePlass)
		{
			result.Add($"{plass.PlassId}: Framleid plass, men ikke merker med \"Framleie\"");
		}

		// Sjekk om eier av plass er merket som sluttet
		if (plass.Eier != null && plass.Eier.Contains("Sluttet"))
		{
			result.Add($"{plass.PlassId}: Eier er merket som sluttet i medlemsregisteret");
		}

		if (plass.TilLeie && plass.Eier == null)
		{
			result.Add($"{plass.PlassId}: Plassen er merket til leie, men har ingen eier");
		}

		if (plass.UngdomsPlass)
		{
			result.Add($"{plass.PlassId}: Ungdomsplass opphører. Varsle {plass.Eier} ");
		}

		var venterPaInnskudd = innskuddVenteliste.FirstOrDefault(v => v.PlassId == plass.PlassId);
		if (venterPaInnskudd != null && venterPaInnskudd.Utbetales)
		{
			if (ledigePlasser.Find(p => p.PlassId == plass.PlassId) == null)
			{
				result.Add($"{plass.PlassId}: Denne plassen skal være ledig inntil {venterPaInnskudd.Navn} får tilbakebetalt innskudd");
			}
		}

		if (plass.AndelsPlass || plass.SesongPlass)
		{
			double bredde = plass.Bredde / 100.0;
			var forventetVariant = VareVariant.Create(bredde);
			if (plass.VareVariant.Gruppe != forventetVariant.Gruppe)
			{
				result.Add($"{plass.PlassId}: Feil båtplass-avgift varevariant. Skal være \"{forventetVariant.Text}\"");
			}
		}
	}

	return result;
}

void PrintVentelister(List<PlassSoker> batplassVenteListe, List<BatPlass> andelsplasser)
{
	Console.WriteLine("\n+Venteliste");
	//Console.WriteLine($"\n*** Venteliste ({batplassVenteListe.Count()}) ***\n");
	Console.WriteLine("Medlem                    Plass      Bredde     Lengde     Båt                  Dager      Gruppe");
	Console.WriteLine("-------------------------------------------------------------------------------------------------------");

	var pri1Plasser = batplassVenteListe.Where(p => p.PlassType == PlassType.AndelBytte).OrderBy(p => p.FraTid);
	PrintVenteliste("Pri 1: Bytte av andelsplass", pri1Plasser);

	var underPri1Plasser = batplassVenteListe.Except(pri1Plasser);
	var pri2Plasser = underPri1Plasser
		.Where(p => p.PlassType == PlassType.AndelNy && andelsplasser.Find(a => a.Eier == p.Navn) != null).OrderBy(p => p.FraTid);
	PrintVenteliste("Pri 2: Ny (ekstra) andelsplass for andelshaver", pri2Plasser);

	var underPri2Plasser = underPri1Plasser.Except(pri2Plasser);
	var pri3Plasser = underPri2Plasser
		.Where(p => p.PlassType == PlassType.AndelNy).OrderBy(p => p.FraTid);
	PrintVenteliste("Pri 3: Ny andelsplass", pri3Plasser);

	var underPri3Plasser = underPri2Plasser.Except(pri3Plasser);
	var pri4Plasser = underPri3Plasser
		.OrderBy(p => p.PlassType).ThenBy(p => p.FraTid);
	PrintVenteliste("Pri 4: Sesongplass", pri4Plasser);

	Console.WriteLine("-");
}

void PublishStatus(string file, bool collapse = false)
{
	var style = new List<string>
	{
		"<style>",
		"body {",
		"  font-family: Arial, Helvetica, sans-serif;",
		"}",
		"",
		"/* Make the main heading look clean */",
		"h1 {",
		"  font-family: Arial, Helvetica, sans-serif;",
		"  font-size: 1.8em;",
		"  margin-bottom: 0.5em;",
		"}",
		"",
		"/* Style the summary text */",
		"summary {",
		"  font-family: Arial, Helvetica, sans-serif;",
		"  font-size: 1.1em;",
		"  cursor: pointer;",
		//"  font-weight: bold;",
		"  list-style: none;",
		"}",
		"",
		"/* Remove default marker and use + / - */",
		"summary::marker {",
		"  display: none;",
		"}",
		"",
		"summary::before {",
		"  content: \"+ \";",
		"  font-weight: bold;",
		"  margin-right: 5px;",
		"}",
		"",
		"details[open] summary::before {",
		"  content: \"-\";",
		"}",
		"</style>",
		"<h1>StyreWeb Status</h1>"
	};

	var prefix = new List<string>
	{
		"<!DOCTYPE html>",
		"<html>",
		"<head>",
		"    <meta charset=\"UTF-8\">",
		"    <title>StyreWeb status</title>",
		"</head>",
		"<body>",
		"<pre>"
	};

	var postfix = new List<string>
	{
		"</pre>",
		"</body>",
		"</html>"
	};

	var contents = File.ReadAllLines(file);
	var path = Path.GetDirectoryName(file);
	var fileName = Path.GetFileNameWithoutExtension(file);
	var htmlFile = Path.Combine(path, fileName + ".html");
	using (var writer = new StreamWriter(htmlFile, false, Encoding.UTF8))
	{
		(collapse ? style : prefix).ForEach(p => writer.WriteLine(p));
		foreach (var line in contents)
		{
			if (line.StartsWith("+"))
			{
				if (collapse)
				{
					writer.WriteLine("<details>");
					writer.WriteLine($"  <summary>{line.Substring(1)}</summary>");
					writer.WriteLine("  <pre>");
				}
				else
				{
					writer.WriteLine(line.Substring(1));	// Skip +
				}
			}
			else if (line == "-")
			{
				if (collapse)
				{
					writer.WriteLine("  </pre>");
					writer.WriteLine("</details>");
					//writer.WriteLine();
				}
			}
			else
			{
				writer.WriteLine(line);
			}
		}
		
		if (!collapse)
		{
			foreach (var line in postfix)
			{
				writer.WriteLine(line);
			}
		}
	}

	UploadFtp($"ftp://vd03.verdidata.no/{fileName}.html", htmlFile, "solvikenftp", "c5712$vQp");
	Console.WriteLine($"Lastet opp {htmlFile} til solviken.no");
}

void UploadFtp(string ftpUrl, string localFile, string user, string pass)
{
	FtpWebRequest req = (FtpWebRequest)WebRequest.Create(ftpUrl);
	req.Method = WebRequestMethods.Ftp.UploadFile;
	req.Credentials = new NetworkCredential(user, pass);
	req.UsePassive = true;
	req.UseBinary = true;
	req.KeepAlive = false;

	byte[] data = File.ReadAllBytes(localFile);
	req.ContentLength = data.Length;

	try
	{
		using (Stream stream = req.GetRequestStream())
			stream.Write(data, 0, data.Length);
	}
	catch (Exception ex)
	{
		return;
	}

	using (FtpWebResponse resp = (FtpWebResponse)req.GetResponse())
		Console.WriteLine(resp.StatusDescription);
}

void PrintVenteliste(string heading, IEnumerable<PlassSoker> venteliste)
{
	Console.WriteLine($"\n* {heading} ({venteliste.Count()}):");
	foreach (var plass in venteliste)
	{
		int dager = (DateTime.Now - plass.FraTid).Days;
		var gruppe = HavneData.FinnInnskuddGruppeFraLengde(plass.Lengde);
		var gruppeTxt = $"{gruppe.Item1}: {gruppe.Item2}";
		Console.WriteLine($"{plass.Navn,-25} {plass.PlassId,-10} {plass.Bredde,-10} {plass.Lengde,-10} {plass.BatType,-20} {dager,-10} {gruppeTxt}");
	}
}

void PrintBatplass(BatPlass plass, bool visEier, MedlemsRegister medlemsRegister, string postfix = null)
{
	var bruker = visEier ? plass.Eier : plass.Leier;
	var bredde = ((double)plass.Bredde / 100).ToString("0.00");
	var lengde = ((double)plass.Lengde / 100).ToString("00.00").TrimStart('0').PadLeft(5);
	
	if (bruker == null)
	{
		Console.WriteLine($"*** Feil - plass {plass.PlassId} har ingen eier");
		return;
	}
	
	if (!medlemsRegister.Medlemmer.TryGetValue(bruker, out var xxx))
	{
		Console.WriteLine($"*** Feil - {bruker} eier ikke plass {plass.PlassId}");
		return;
	}
	
	var eier = TilpassNavn(bruker);
	Console.Write($"{plass.PlassId.Substring(0, 4)}: {eier,-25} BxL: {bredde} x {lengde}");
	
	if (postfix != null)
	{
		Console.WriteLine($"  {postfix}");
	}
	else
	{
		Console.WriteLine();
	}
}

void PrintBatplass2(BatPlass plass, MedlemsRegister medlemsRegister)
{
	var bredde = ((double)plass.Bredde / 100).ToString("0.00");
	var lengde = ((double)plass.Lengde / 100).ToString("00.00").TrimStart('0').PadLeft(5);
	var bruker = plass.Bruker;
	if (bruker == null)
	{
		if (plass.Reservert)
		{
			// Reservert plass
			Console.WriteLine($"{plass.PlassId.Substring(0, 4)}: *** Reservert ***");
		}
		else
		{
			// Ledig plass
			Console.WriteLine($"{plass.PlassId.Substring(0, 4)}: *** Ledig ***");
		}
		return;
	}

	string postfix;
	if (plass.UngdomsPlass)
	{
		postfix = "(Ungdomsplass)";
	}
	else if (plass.JollePlass)
	{
		postfix = "(Jolleplass)";
	}
	else if (plass.TilLeie)
	{
		postfix = "(Til leie)";
	}
	else
	{
		if (plass.Leier != null && plass.Eier != null)
		{
			if (plass.SesongPlass)
			{
				postfix = $"(Utleie fra {plass.Eier})";
			}
			else if (plass.LanePlass)
			{
				postfix = $"(Lån fra {plass.Eier})";
			}
			else
			{
				postfix = null;
			}
		}
		else
		{
			postfix = null;
		}
	}
	
	var eier = TilpassNavn(bruker);
	Console.Write($"{plass.PlassId.Substring(0, 4)}: {eier,-25} BxL: {bredde} x {lengde}");
	
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
	public DateTime UtlevertFra { get; set; }
	public DateTime UtLeidFra { get; set; }
	public string Leier { get; set; }
	public int Bredde { get; set; }
	public int Lengde { get; set; }
	public char Gruppe { get; set; }
	public int Innskudd { get; set; }
	public bool AndelsPlass { get; set; }
	public bool SesongPlass { get; set; }
	public bool FramleiePlass { get; set; }
	public bool UngdomsPlass { get; set; }
	public bool JollePlass { get; set; }
	public bool LanePlass { get; set; }
	public bool TilLeie { get; set; }
	public bool Reservert { get; set; }
	public bool LandOpplag { get; set; }
	public string Vaktfritak { get; set; }
	public string batType { get; set; }
	public VareVariant VareVariant { get; set; }
	public string Bruker => Leier ?? Eier; 

	List<(double limit, int factor)[]> prisgrupper = Priser.Prisgrupper;
	
	public int GetPrisGruppe()
	{
		double bredde = (double)Bredde / 100;
		for (int i = 0; i < prisgrupper[0].Length; i++)
		{
			var prisgruppe = prisgrupper[0][i];
			if (bredde < prisgruppe.limit)
			{
				return i;
			}
		}
		
		return -1;
	}
	
	public int GetPrisFaktor(int prismodell)
	{
		int gruppe = GetPrisGruppe();
		return prisgrupper[prismodell][gruppe].factor;
	}
	
	public int BeregnBatplassAvgift()
	{
		if (UngdomsPlass)
		{
			return 1000;
		}

		double bredde = (double)Bredde / 100;
		double lengde = (double)Lengde / 100;
		int beregnetAvgift = (int)Math.Round(bredde * lengde * PrisFaktor(lengde));
		if (beregnetAvgift < 2500)
		{
			beregnetAvgift = 2500;
		}

		return beregnetAvgift;
		
		//var leiePlass = Leier != null;
		//return beregnetAvgift + (leiePlass ? LeieTillegg(lengde) : 0);
	}
	
	public int BeregnNyBatplassAvgift(int variant, bool includeLimit = false)
	{
		double bredde = (double)Bredde / 100;
		int beregnetAvgift = (int)Math.Round(bredde * NyPrisFaktor(variant, includeLimit, bredde));
		if (beregnetAvgift < 2500)
		{
			beregnetAvgift = 2500;
		}
		
		return beregnetAvgift;
	}

	public int PrisFaktor(double lengde)
	{
		return lengde <= 7.2 ? 160 : (lengde >= 9.2 ? 200 : 180);
	}

	public int NyPrisFaktor(int variant, bool includeLimit, double bredde)
	{
		if (variant < prisgrupper.Count)
		{
			foreach (var prisgruppe in prisgrupper[variant])
			{
				if (includeLimit)
				{
					if (bredde <= prisgruppe.limit)
					{
						return prisgruppe.factor;
					}
				}
				else
				{
					if (bredde < prisgruppe.limit)
					{
						return prisgruppe.factor;
					}
				}
			}
		}
		return 0;
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
	protected SortedDictionary<string, BatPlass> BatPlasser { get; set; }
	
	public abstract string Navn { get; }
	
	public DateTime TimeStamp { get; set; }
	
	public string PlassPrefix { get; set; }
	
	public List<PlassSoker> BatplassVenteliste { get; set; }
	
	public List<InnskuddEier> InnskuddVenteliste { get; set; }

	protected abstract HavneData Read(string fromDate = null);
	
	public HavneData LesData(string prefix = null, string fromDate = null)
	{
		PlassPrefix = prefix;
		Read(fromDate);
		
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

	public List<BatPlass> GetAlleBryggePlasser()
	{
		return BatPlasser.Values.Where(p => !p.LandOpplag).ToList();
	}

	public List<BatPlass> GetAndelsPlasser()
	{
		return BatPlasser.Values.Where(v => v.AndelsPlass).ToList();
	}

	public List<BatPlass> GetSesongPlasser()
	{
		return BatPlasser.Values.Where(v => v.SesongPlass).ToList();
	}

	public List<BatPlass> GetFramleiePlasser()
	{
		return BatPlasser.Values.Where(v => v.FramleiePlass).ToList();
	}

	public List<BatPlass> GetUngdomsPlasser()
	{
		return BatPlasser.Values.Where(v => v.UngdomsPlass).ToList();
	}

	public List<BatPlass> GetJollePlasser()
	{
		return BatPlasser.Values.Where(v => v.JollePlass).ToList();
	}

	public List<BatPlass> GetLanePlasser()
	{
		return BatPlasser.Values.Where(v => v.LanePlass).ToList();
	}

	public List<BatPlass> GetTilLeiePlasser()
	{
		return BatPlasser.Values.Where(v => v.TilLeie).ToList();
	}

	public List<BatPlass> GetReservertePlasser()
	{
		return BatPlasser.Values.Where(v => v.Reservert).ToList();
	}
	
	public List<BatPlass> GetLandOpplagsPlasser()
	{
		return BatPlasser.Values.Where(v => v.LandOpplag).ToList();
	}

	public List<BatPlass> GetLedigePlasser()
	{
		return BatPlasser.Values
		.Except(GetLandOpplagsPlasser())
		.Except(GetAndelsPlasser())
		.Except(GetSesongPlasser())
		.Except(GetUngdomsPlasser())
		.Except(GetJollePlasser())
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
			double lengde = (double)plass.Lengde / 100;
			var gruppe = FinnInnskuddGruppeFraLengde(lengde);
			return (gruppe.Item3, gruppe.Item1);
		}
		
		return (-1, "X");
	}
	
	public static (string, string, int) FinnInnskuddGruppeFraLengde(double lengde)
	{
		switch (lengde)
		{
			case double len when (len <= 5.4):
				return ("A", "Inntil 5,4 m", 7750);
			case double len when (len <= 7.0):
				return ("B", "5,5 - 7,0 m", 11100);
			case double len when (len <= 8.7):
				return ("C", "7,1 – 8,7 m", 14700);
			case double len when (len <= 9.1):
				return ("D", "8,8 – 9,1 m", 19750);
			case double len when (len <= 10.0):
				return ("E", "9,2 – 10,0 m", 22500);
			case double len when (len <= 10.6):
				return ("F", "10,1 – 10,6 m", 25150);
			case double len when (len <= 11.8):
				return ("G", "10,7 – 11,8 m", 27900);
			case double len when (len <= 12.4):
				return ("H", "11,9 – 12,4 m", 34500);
			default:
				return ("L", "12,5 – 13,7 m", 45500);
		}
	}
}

public class StyreWebExport : HavneData
{
	private string downloadFolder;
	private string swMarinaDetaljertFil;
	private string swFramleieFil;
	private string swGruppeBatplassVenteliste;
	private string swGruppeInnskuddVenteliste;
	private string swExportFolder;
	private List<string> swFritaksGrupper;

	public override string Navn { get => "StyreWeb"; }

	public StyreWebExport()
	{
		downloadFolder = $@"C:\Users\{Environment.UserName}\Downloads";
		var workFolder = @"C:\MyLocal\Solviken";
		swExportFolder = Path.Combine(workFolder, "FraStyreweb");
		swGruppeBatplassVenteliste = "Venteliste";
		swGruppeInnskuddVenteliste = "Innskudd_uten_båt";
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

	protected override HavneData Read(string fromDate = null)
	{
		if (fromDate == null)
		{
			CopyNewerFile(Path.Combine(downloadFolder, "Marina_-_Detaljert.csv"), swExportFolder);
			CopyNewerFile(Path.Combine(downloadFolder, "Fremleie_historie.csv"), swExportFolder);

			foreach (var gruppe in swFritaksGrupper
								.Append(swGruppeBatplassVenteliste)
								.Append(swGruppeInnskuddVenteliste))
			{
				CopyNewerFile(Path.Combine(downloadFolder, $"Gruppe{gruppe}.xlsx"), swExportFolder);
				ConvertFromXlsx2Csv(Path.Combine(swExportFolder, $"Gruppe{gruppe}.xlsx"));
			}

			CopyNewerFile(Path.Combine(downloadFolder, "Dugnadsregnskap.xlsx"), swExportFolder);
			ConvertFromXlsx2Csv(Path.Combine(swExportFolder, "Dugnadsregnskap.xlsx"));
		}
		else
		{
			if (DateTime.TryParse(fromDate, out var from))
			{
				var subFolders = Directory.GetDirectories(swExportFolder);
				var backups = new List<DateTime>();
				foreach (var folder in subFolders)
				{
					var levels = folder.Split('\\');
					var dateString = levels[levels.Length - 1];
					if (DateTime.TryParse(dateString, out var date))
					{
						backups.Add(date);
					}
				}
				
				if (backups.Count > 0)
				{
					backups.Sort();
					string date = null;
					for (int i = backups.Count - 1; i >= 0; i--)
					{
						if (backups[i] <= from)
						{
							date = backups[i].ToString("d");
							swExportFolder = Path.Combine(swExportFolder, date);
							Console.WriteLine($"Leser Styreweb data fra {date}");
							break;
						}
					}
					
					if (date == null)
					{
						date = backups[0].ToString("d");
						swExportFolder = Path.Combine(swExportFolder, date);
						Console.WriteLine($"Fant ikke StyreWeb data for {fromDate}, henter fra eldste backup: {date}");
					}
				}
			}
		}

		swMarinaDetaljertFil = Path.Combine(swExportFolder, "Marina_-_Detaljert.csv");
		swFramleieFil = Path.Combine(swExportFolder, "Fremleie_historie.csv");
		
		if (File.Exists(swMarinaDetaljertFil))
		{
			TimeStamp = File.GetLastWriteTime(swMarinaDetaljertFil);
			using (var reader = new StreamReader(swMarinaDetaljertFil, Encoding.GetEncoding("UTF-8")))
			{
				reader.ReadLine();      // Skip header
				string line;
				while ((line = reader.ReadLine()) != null && !line.StartsWith("#"))
				{
					var fields = line.Split('\t');
					var plassId = fields[1];
					var plassType = fields[2];
					var breddeMeter = fields[3];
					var lengdeMeter = fields[4];
					var gruppe = fields[7];
					var innskudd = fields[8];
					DateTime utlevertFra;
					if (DateTime.TryParse(fields[13], out var time))
					{
						utlevertFra = time;
					}
					else
					{
						utlevertFra = DateTime.Now;
					}
					var eier = fields[14];
					var vareVariant = VareVariant.Create(fields[19]);

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
					var framleiePlass = (plassType == "Framleie");
					var ungdomsPlass = (plassType == "Ungdomsplass");		// Denne skal vekk i 2026
					var jollePlass = (plassType == "Jolleplass");
					var lanePlass = (plassType == "Låneplass");				// Framleie for en periode, eier betaler, framleier ikke
					var reservert = (plassType == "Reservert");
					var tilLeie = (plassType == "Til leie");
					var landOpplag = (plassType == "Landopplag");
					var andelsPlass = //((eier != null) && !landOpplag) ||	// Gammel definisjon
										(plassType == "Andelsplass") || framleiePlass || lanePlass || tilLeie;

					int innskuddKr = 0;
					if (innskudd.Length > 0)
					{
						innskuddKr = int.Parse(innskudd.Split(',')[0]);
					}
					
					BatPlasser[plassId] = new BatPlass
					{
						PlassId = plassId,
						Eier = eier,
						UtlevertFra = utlevertFra,
						Bredde = breddeCm,
						Lengde = lengdeCm,
						AndelsPlass = andelsPlass,
						SesongPlass = sesongPlass,
						FramleiePlass = framleiePlass,
						UngdomsPlass = ungdomsPlass,
						JollePlass = jollePlass,
						LanePlass = lanePlass,
						TilLeie = tilLeie,
						Reservert = reservert,
						LandOpplag = landOpplag,
						VareVariant = vareVariant,
						Gruppe = gruppe.Length > 0 ? gruppe[0] : ' ',
						Innskudd = innskuddKr
					};
				}
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
				var utleidFra = fields[2];
				var tilDato = fields[3];
				var leier = fields[4];
				
				if (leier != "")
				{
					if (BatPlasser.TryGetValue(plassId, out var batPlass))
					{
						if (DateTime.TryParse(tilDato, out var sluttDato))
						{
							if (sluttDato > DateTime.Now)
							{
								batPlass.Leier = leier;
								
								if (DateTime.TryParse(utleidFra, out var startDato))
								{
									batPlass.UtLeidFra = startDato;
								}

								if (!(batPlass.SesongPlass
										|| batPlass.FramleiePlass
										|| batPlass.UngdomsPlass
										|| batPlass.JollePlass
										|| batPlass.LanePlass))
								{
									Console.WriteLine($"Plass {plassId} framleid, feil plasstype");
								}
							}
						}
						else
						{
							Console.WriteLine($"Plass {plassId} framleie, mangler sluttdato");
						}
					}
					else
					{
						//Console.WriteLine($"Plass \"{plassId}\" framleid, finnes ikke i marina");
					}
				}
				else
				{
					Console.WriteLine($"Plass {plassId} framleie uten leietager");
				}
			}
		}

		var kjenteFritak = swFritaksGrupper
		.SelectMany(LesGruppe)
		.OrderBy(g => g.Item1)
		.Distinct(new VervNavnComparer())
		.ToDictionary(key => key.Item1, value => value.Item2);
		var andelsplasser = GetAndelsPlasser();
		var sesongplasser = GetSesongPlasser();
		var pliktigePlasser = andelsplasser.Concat(sesongplasser).OrderBy(a => a.PlassId);
		var pliktige = pliktigePlasser.Select(p => p.Bruker).Distinct();
		var fritak = pliktigePlasser.Where(p => kjenteFritak.TryGetValue(p.Bruker, out var foo));
		
		foreach (var plass in fritak)
		{
			plass.Vaktfritak = kjenteFritak[plass.Bruker];
			//var bruker = plass.Bruker;
			//if (bruker != null)
			//{
			//	if (kjenteFritak.TryGetValue(bruker, out var reason))
			//	{
			//		plass.Vaktfritak = reason;
			//	}
			//	else
			//	{
			//		plass.Vaktfritak = "Fritak, ukjent årsak";
			//	}
			//}
		}

		BatplassVenteliste = LesBatplassVenteliste();
		InnskuddVenteliste = LesInnskuddVenteliste();
		
		return this;
	}

	private List<(string, string)> LesGruppe(string gruppe)
	{
		var gruppeFil = Path.Combine(swExportFolder, $"Gruppe{gruppe}.csv");
		var medlemmer = new List<(string, string)>();   // (Navn, gruppe)
		if (!File.Exists(gruppeFil) && !swExportFolder.EndsWith("FraStyreweb"))
		{
			// Try one level up instead
			var folder = swExportFolder + @"\..";
			gruppeFil = Path.Combine(folder, $"Gruppe{gruppe}.csv");
		}
		if (File.Exists(gruppeFil))
		{
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
			}
		}

		return medlemmer;
	}

	private List<PlassSoker> LesBatplassVenteliste()
	{
		var gruppeFil = Path.Combine(swExportFolder, "GruppeVenteliste.csv");
		var sokere = new List<PlassSoker>();
		if (File.Exists(gruppeFil))
		{
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
					var dato = fields[10];
					if (!ParseDate(dato, out var fraTid))
					{
						Console.WriteLine($"Søker {navn} har ugyldig starttid {dato}");
						continue;
					}

					var kommentar = fields[12].Trim('"');
					var felt = kommentar.Split(';');
					if (felt.Length != 6)
					{
						Console.WriteLine($"Søker {navn} har ugyldig beskrivelse (\"kommentar\") {kommentar}");
						continue;
					}

					sokere.Add(
						new PlassSoker
						{
							Navn = navn,
							FraTid = fraTid,
							PlassType = GetPlassType(felt[0], felt[1]),
							PlassId = GetPlassId(felt[1]),
							Seilbat = (felt[2] == "S"),
							Bredde = double.Parse(felt[3]),
							Lengde = double.Parse(felt[4]),
							BatType = felt[5]
						}
					);
				}
			}
		}
		
		return sokere;
	}

	private bool ParseDate(string date, out DateTime tid)
	{
		return DateTime.TryParseExact(date, "M/d/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out tid)
		|| DateTime.TryParse(date, out tid);
	}

	private List<InnskuddEier> LesInnskuddVenteliste()
	{
		var gruppeFil = Path.Combine(swExportFolder, "GruppeInnskudd_uten_båt.csv");
		var ventere = new List<InnskuddEier>();
		if (File.Exists(gruppeFil))
		{
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
					var kommentar = fields[12].Trim('"');
					var felt = kommentar.Split(';');
					if (felt.Length != 3)
					{
						Console.WriteLine($"Innskuddeier {navn} har ugyldig beskrivelse (\"kommentar\") {kommentar}");
						continue;
					}

					ventere.Add(
						new InnskuddEier
						{
							Navn = navn,
							PlassId = felt[0],
							Innskudd = int.Parse(felt[1]),
							Utbetales = felt[2][0] == 'U'
						}
					);
				}
			}
		}

		return ventere;
	}

	private PlassType GetPlassType(string v1, string v2)
	{
		switch (v1)
		{
			case "A":
				// Andelsplass
				switch (v2[0])
				{
					case 'B':
						return PlassType.AndelBytte;
					case 'N':
						return PlassType.AndelNy;
				}
			break;
			case "S":
				switch (v2[0])
				{
					case 'F':
						return PlassType.SesongForny;
					case 'N':
						return PlassType.SesongNy;
				}
			break;
			case "J":
				return PlassType.Jolle;
		}
		
		return PlassType.Ingen;
	}

	private string GetPlassId(string v2)
	{
		if (v2.Length == 5)
		{
			return v2.Substring(1);
		}
		
		return null;
	}

	private void ConvertFromXlsx2Csv(string excelFile)
	{
		if (!File.Exists(excelFile))
		{
			return;
		}
		
		using (var stream = File.Open(excelFile, FileMode.Open, FileAccess.Read))
		{
			using (var reader = ExcelReaderFactory.CreateReader(stream))
			{
				var path = Path.GetDirectoryName(excelFile);
				var fileName = Path.GetFileNameWithoutExtension(excelFile);
				var csvFile = Path.Combine(path, fileName) + ".csv";
				using (var writer = new StreamWriter(csvFile))
				{
					var result = reader.AsDataSet();

					DataTable table = result.Tables[0]; // first sheet

					foreach (DataRow row in table.Rows)
					{
						for (int i = 0; i < row.ItemArray.Length - 1; i++)
						{
							writer.Write($"{row.ItemArray[i]}\t");
						}
						writer.WriteLine($"{row.ItemArray[row.ItemArray.Length - 1]}");
					}
				}
			}
		}
	}

	private void CopyNewerFile(string source, string destination)
	{
		var fileName = Path.GetFileName(source);
		var destinationFile = Path.Combine(destination, fileName);
		
		if (File.Exists(source))
		{
			if (File.Exists(destinationFile))
			{
				File.Delete(destinationFile);
			}
			
			File.Move(source, destinationFile);
			Console.WriteLine($"Oppdaterte StyreWeb export fil \"{fileName}\" fra Nedlastinger");
			
			if (!fileName.StartsWith("Gruppe"))		// Grupperingene er (stort sett) faste
			{
				var date = DateTime.Now.ToString("d");
				var backupPath = Path.Combine(destination, date);
				Directory.CreateDirectory(backupPath);
				File.Copy(destinationFile, Path.Combine(backupPath, fileName), true);
			}
		}
	}
}

public class PlassSoker
{
	public string Navn { get; set; }
	public DateTime FraTid { get; set; }
	public PlassType PlassType { get; set; }
	public string PlassId { get; set; }		// Hvis A;B, A;S eller S;F
	public bool Seilbat { get; set; }
	public double Bredde { get; set; }
	public double Lengde { get; set; }
	public string BatType { get; set; }
}

public enum PlassType
{
	AndelBytte,
	AndelNy,
	SesongForny,
	SesongNy,
	Jolle,
	Ingen
}

public class InnskuddEier
{
	public string Navn { get; set; }
	public string PlassId { get; set; }
	public int Innskudd { get; set; }
	public bool Utbetales { get; set; }
}

public class ExcelExport : HavneData
{
	private string batplassFil;

	public override string Navn { get => "Excel"; }

	public ExcelExport()
	{
		var workFolder = @"C:\MyLocal\Solviken";
		batplassFil = Path.Combine(workFolder, "Batplasser.csv");
		BatPlasser = new SortedDictionary<string, BatPlass>();
	}

	protected override HavneData Read(string fromDate = null)
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
					Bredde = breddeCm,
					Lengde = lengdeCm
				};
			}
		}

		return this;
	}
}

public class HavneWebExport : HavneData
{
	private string hwExportFil;

	public override string Navn { get => "HavneWeb"; }

	public HavneWebExport()
	{
		var workFolder = @"C:\Users\Solviken\OneDrive\Solviken\2025\Havnedatabasen";
		hwExportFil = Path.Combine(workFolder, "SolvikenBtforening_311224_124226.csv");
		BatPlasser = new SortedDictionary<string, BatPlass>();
	}

	protected override HavneData Read(string fromDate = null)
	{
		TimeStamp = DateTime.Parse("31.12.2024");
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
					var gruppe = fields[3];
					var breddeMeter = fields[4];
					var lengdeMeter = fields[5];
					var innskudd = fields[9];
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
						Gruppe = gruppe.Length > 0 ? gruppe[0] : ' ',
						Eier = eier,
						Innskudd = innskuddKr,
						Leier = leier,
						Bredde = breddeCm,
						Lengde = lengdeCm,
						Vaktfritak = vaktFritak,
						batType = batType
					};
				}
			}
		}

		return this;
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
	public string Avdeling { get; set; }
}

public class VareVariant
{
	private static (double limit, string text)[] varevarianter = new (double limit, string text)[]
	{
		(2.5, "Bredde < 2,5 m"),
		(3.0, "Bredde 2,5 - 2,99 m"),
		(3.5, "Bredde 3,0 - 3,49 m"),
		(4.0, "Bredde 3,5 - 3,99 m"),
		(4.5, "Bredde 4,0 - 4,49 m"),
		(5.0, "Bredde 4,5 - 4,99 m"),
		(9.0, "Bredde >= 5,0 m")
	};
	
	private VareVariant(int gruppe)
	{
		Gruppe = gruppe;
		Text = varevarianter[gruppe].text;
	}
	
	public int Gruppe { get; }		// 0-6
	
	public string Text { get; }
	
	public static VareVariant Create(double bredde)
	{
		for (int i = 0; i < varevarianter.Length; i++)
		{
			if (bredde < varevarianter[i].limit)
			{
				return new VareVariant(i);
			}
		}
		
		return new VareVariant(0);
	}
	
	public static VareVariant Create(string variantText)
	{
		for (int i = 0; i < varevarianter.Length; i++)
		{
			if (variantText == varevarianter[i].text)
			{
				return new VareVariant(i);
			}
		}
		
		return new VareVariant(0);
	}
}

public class MedlemsRegister
{
	private string swFolder;
	private string swEksportFil;
	private string downloadFolder;

	public Dictionary<string, Medlem> Medlemmer { get; private set; }
	
	public MedlemsRegister()
	{
		swFolder = @"C:\MyLocal\Solviken\FraStyreWeb";
		downloadFolder = $@"C:\Users\{Environment.UserName}\Downloads";
		swEksportFil = "Detaljert_Rapport.csv";
	}
	
	public MedlemsRegister LesData()
	{
		Medlemmer = new Dictionary<string, Medlem>();
		CopyNewerFile(Path.Combine(downloadFolder, swEksportFil), swFolder);
		using (var reader = new StreamReader(Path.Combine(swFolder, swEksportFil), Encoding.GetEncoding("UTF-8")))
		{
			reader.ReadLine();      // Skip header
			string line;
			while ((line = reader.ReadLine()) != null)
			{
				var fields = line.Split('\t');
				if (fields.Length >= 5)
				{
					var name = $"{fields[1]} {fields[0]}";
					var avd = fields[4];
					var tlf = fields[16];
					var epost = fields[17];
					Medlemmer[name] = new Medlem {Navn = name, Avdeling = avd, Tlf = tlf, Epost = epost};
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

public static class Priser
{
	public static List<(double limit, int factor)[]> Prisgrupper = new List<(double limit, int factor)[]>
	{
		new[] { (2.5, 800), (3.0, 1005), (3.5, 1050), (4.0, 1420), (4.5, 2000), (5.0, 2200), (10.0, 2500) },	// Joakims tall
		new[] { (2.5, 1110), (3.0, 1012), (3.5, 1065), (4.0, 1405), (4.5, 1625), (5.0, 1895), (10.0, 1810) },	// Min. faktor for hver gruppe
		new[] { (2.5, 1000), (3.0, 1100), (3.5, 1200), (4.0, 1400), (4.5, 1600), (5.0, 1800), (10.0, 2000) },	// Glattet ut
		new[] { (2.5, 1000), (3.0, 1050), (3.5, 1100), (4.0, 1450), (4.5, 1700), (5.0, 1900), (10.0, 2000) },	// Glattet ut
	};
}