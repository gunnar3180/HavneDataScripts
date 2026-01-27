<Query Kind="Program">
  <Reference>&lt;RuntimeDirectory&gt;\System.Collections.Concurrent.dll</Reference>
  <Namespace>System.Collections.Concurrent</Namespace>
  <Namespace>System.Globalization</Namespace>
</Query>

void Main()
{
	bool bryggeliste = false;
	//
	//
	//Enumerable.Range(1, 6)
	//	.ToList()
	//	.ForEach(e => VisAlledata(new StyreWebExport().LesData(e.ToString()), true, @"C:\MyLocal\Solviken\Rapporter"));
	
	VisAlleData(new StyreWebExport().LesData(), bryggeliste);
	//VisAlledata(new ExcelExport().LesData(), bryggeliste);
	//VisAlledata(new HavneWebExport().LesData(), bryggeliste);

	//VisEierEndringer(new StyreWebExport().LesData(fromDate: "21.08.2025"),
	//				 new StyreWebExport().LesData());
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
	//SjekkVareVarianter(new StyreWebExport().LesData());
	//FinnPlasserUnder2500(new StyreWebExport().LesData());
	//FinnEierEndringerEtter(DateTime.Parse("20.06.2025"), new StyreWebExport().LesData());
	//FinnSesongLeiereFraAndelsplass(new StyreWebExport().LesData());
	//SammenlignGrupper(new HavneWebExport().LesData(), new StyreWebExport().LesData());
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
		var lengde = ((double)plass.Lengde) / 100;
		var beregnetVareVariant = VareVariant.Create(lengde);
		if (plass.VareVariant.Size != beregnetVareVariant.Size)
		{
			Console.WriteLine($"{plass.PlassId}: Feil varevariant {plass.VareVariant.Size}, skal være {beregnetVareVariant.Size}");
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

void VisAlleData(HavneData dataSet, bool bryggeliste = true, string path = null)
{
	var andelsplasser = dataSet.GetAndelsPlasser();
	var sesongplasser = dataSet.GetSesongPlasser();
	var ungdomsplasser = dataSet.GetUngdomsPlasser();
	var jollePlasser = dataSet.GetJollePlasser();
	var tilLeiePlasser = dataSet.GetTilLeiePlasser();
	var ledigePlasser = dataSet.GetLedigePlasser();
	var venteListe = dataSet.VenteListe;
	var medlemsRegister = new MedlemsRegister().LesData();
	StreamWriter writer = null;
	
	if (path != null)
	{
		writer = new StreamWriter(Path.Combine(path, $"Brygge{dataSet.PlassPrefix ?? "r"}.txt"), false, Encoding.GetEncoding("UTF-8"));
		Console.SetOut(writer);
	}
	
	Console.Write($"Eksport fra {dataSet.Navn}");
	if (dataSet.PlassPrefix != null)
	{
		Console.WriteLine($" - Plasser som starter med \"{dataSet.PlassPrefix}\"");
	}
	else
	{
		Console.WriteLine();
	}

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
		Console.WriteLine($"{plass.PlassId}: {plass.Leier,-25} {tlf,-13}");
	}

	Console.WriteLine($"\n{jollePlasser.Count} jolleplasser");
	foreach (var plass in jollePlasser)
	{
		var tlf = medlemsRegister.Medlemmer[plass.Leier].Tlf;
		Console.WriteLine($"{plass.PlassId}: {plass.Leier,-25} {tlf,-13}");
	}

	// Til leie-plasser
	Console.WriteLine($"\n{tilLeiePlasser.Count} plasser til leie");
	foreach (var plass in tilLeiePlasser)
	{
		PrintBatplass(plass, true, medlemsRegister);
	}

	Console.WriteLine($"\n{ledigePlasser.Count} ledige plasser");
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
	
	Console.WriteLine("\nLedige plasser pr. gruppe:");
	foreach (var ledige in ledigePrGruppe)
	{
		Console.WriteLine($"{ledige.Key}: {ledige.Value}");
	}
	
	var sortert = breddeListe.OrderBy(l => l.Item2);
	
	Console.WriteLine("\nLedige plasser sortert på bredde:");
	foreach (var plass in sortert)
	{
		Console.WriteLine($"{plass.Item1}: {plass.Item2, 10} m     Gruppe {plass.Item3}");
	}

	Console.WriteLine($"\n *** Venteliste ({venteListe.Count()}) ***\n");
	Console.WriteLine("Medlem                    Plass      Bredde     Lengde     Båt                  Dager      Gruppe");
	Console.WriteLine("-------------------------------------------------------------------------------------------------------");
	
	var pri1Plasser = venteListe.Where(p => p.PlassType == PlassType.AndelBytte).OrderBy(p => p.FraTid);
	PrintVenteliste("Pri 1: Bytte av andelsplass", pri1Plasser);

	var underPri1Plasser = venteListe.Except(pri1Plasser);
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

	Console.WriteLine();
	
	if (writer != null)
	{
		writer.Close();
	}
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
	
	var tlf = medlemsRegister.Medlemmer[bruker].Tlf;
	var eier = TilpassNavn(bruker);
	
	Console.Write($"{plass.PlassId.Substring(0, 4)}: {eier,-25} {tlf,-13} BxL: {bredde} x {lengde}");
	
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
	var bruker = plass.Leier ?? plass.Eier;
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
	
	var tlf = medlemsRegister.Medlemmer[bruker].Tlf;
	var eier = TilpassNavn(bruker);
	Console.Write($"{plass.PlassId.Substring(0, 4)}: {eier,-25} {tlf,-13} BxL: {bredde} x {lengde}");
	
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
	public bool SesongPlass { get; set; }
	public bool UngdomsPlass { get; set; }
	public bool JollePlass { get; set; }
	public bool LanePlass { get; set; }
	public bool TilLeie { get; set; }
	public bool Reservert { get; set; }
	public bool LandOpplag { get; set; }
	public string Vaktfritak { get; set; }
	public string batType { get; set; }
	public VareVariant VareVariant { get; set; }

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

		var leiePlass = Leier != null;
		return beregnetAvgift + (leiePlass ? LeieTillegg(lengde) : 0);
	}

	public int PrisFaktor(double lengde)
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
	protected SortedDictionary<string, BatPlass> BatPlasser { get; set; }
	
	public abstract string Navn { get; }
	
	public string PlassPrefix { get; set; }
	
	public List<PlassSoker> VenteListe { get; set; }

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
	
	public List<BatPlass> GetAndelsPlasser()
	{
		return BatPlasser.Values.Where(v => v.Eier != null)
		.Except(GetLandOpplagsPlasser())
		.ToList();
	}
	
	public List<BatPlass> GetSesongPlasser()
	{
		return BatPlasser.Values.Where(v => v.SesongPlass).ToList();
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
	private string swGruppeVaktplikt;
	private string swGruppeVenteliste;
	private string swExportFolder;
	private List<string> swFritaksGrupper;

	public override string Navn { get => "StyreWeb"; }

	public StyreWebExport()
	{
		downloadFolder = $@"C:\Users\{Environment.UserName}\Downloads";
		var workFolder = @"C:\MyLocal\Solviken";
		swExportFolder = Path.Combine(workFolder, "FraStyreweb");
		swGruppeVaktplikt = "Vaktplikt-2025";
		swGruppeVenteliste = "Venteliste";
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

			foreach (var gruppe in swFritaksGrupper.Append(swGruppeVaktplikt).Append(swGruppeVenteliste))
			{
				CopyNewerFile(Path.Combine(downloadFolder, $"Gruppe{gruppe}.xlsx"), swExportFolder);
				ConvertFromXlsx2Csv(Path.Combine(swExportFolder, $"Gruppe{gruppe}.xlsx"));
			}
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
					var ungdomsPlass = (plassType == "Ungdomsplass");
					var jollePlass = (plassType == "Jolleplass");
					var lanePlass = (plassType == "Låneplass");
					var reservert = (plassType == "Reservert");
					var tilLeie = (plassType == "Til leie");
					var landOpplag = (plassType == "Landopplag");
					
					int innskuddKr = -1;
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
						SesongPlass = sesongPlass,
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
		.SelectMany(g => LesGruppe(g))
		.OrderBy(g => g.Item1)
		.Distinct(new VervNavnComparer())
		.ToDictionary(key => key.Item1, value => value.Item2);
		var pliktigePlasser = GetAndelsPlasser().Concat(GetSesongPlasser());
		var vaktpliktige = LesGruppe(swGruppeVaktplikt);
		var fritak = pliktigePlasser.Where(p => !vaktpliktige.Any(v => (p.Leier??p.Eier) == v.Item1));
		
		foreach (var plass in fritak)
		{
			var bruker = plass.Leier??plass.Eier;
			if (bruker == null)
			{
				
			}
			if (kjenteFritak.TryGetValue(plass.Leier??plass.Eier, out var reason))
			{
				plass.Vaktfritak = reason;
			}
			else
			{
				plass.Vaktfritak = "Fritak, ukjent årsak";
			}
		}

		VenteListe = LesVenteliste();
		
		return this;
	}

	private List<(string, string)> LesGruppe(string gruppe)
	{
		var gruppeFil = Path.Combine(swExportFolder, $"Gruppe{gruppe}.csv");
		var medlemmer = new List<(string, string)>();	// (Navn, gruppe)
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

	private List<PlassSoker> LesVenteliste()
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
					if (!DateTime.TryParseExact(dato, "M/d/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out var fraTid))
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
		var folder = Path.GetDirectoryName(excelFile);
		var csvFile = Path.Combine(folder, Path.GetFileNameWithoutExtension(excelFile)) + ".csv";

		if (File.GetLastWriteTime(excelFile) > File.GetLastWriteTime(csvFile))
		{
			File.Delete(csvFile);
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
			if (File.Exists(destinationFile))
			{
				File.Delete(destinationFile);
			}
			
			File.Move(source, destinationFile);
			Console.WriteLine($"Oppdaterte StyreWeb export fil \"{fileName}\" fra Nedlastinger");
			
			var date = DateTime.Now.ToString("d");
			var backupPath = Path.Combine(destination, date);
			Directory.CreateDirectory(backupPath);
			File.Copy(destinationFile, Path.Combine(backupPath, fileName), true);
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
	public enum Length
	{
		Undefined,
		Small,
		Medium,
		Big
	}
	
	public Length Size { get; set; }
	
	public static VareVariant Create(double length)
	{
		if (length < 7.3)
		{
			return new VareVariant { Size = Length.Small };
		}
		
		if (length < 9.2)
		{
			return new VareVariant { Size = Length.Medium };
		}
		
		return new VareVariant { Size = Length.Big };
	}
	
	public static VareVariant Create(string variantText)
	{
		switch (variantText)
		{
			case "> 7,2 m båtlengde":
				return new VareVariant { Size = Length.Small };
			case "7,3m-9,1m båtlengde":
				return new VareVariant { Size = Length.Medium };
			case "9,2 < båtlengde":
				return new VareVariant { Size = Length.Big };
			default:
				return new VareVariant { Size = Length.Undefined };
		}
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