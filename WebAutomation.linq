<Query Kind="Program">
  <NuGetReference>Microsoft.Playwright</NuGetReference>
  <Namespace>System.Threading.Tasks</Namespace>
  <Namespace>Microsoft.Playwright</Namespace>
</Query>

static async Task Main()
{
	//Microsoft.Playwright.Program.Main(new[] { "install" });   // Install Playwright if not already installed
	await Go().ConfigureAwait(false);
}

public static async Task Go()
{
	var brukerNavn = "gb3180@online.no";
	var kodetPassord = "Cf77D57G1vfiv1S";
	var loginUrl = "https://portal.styreweb.com/account/login.aspx";
	var baseUrl = "https://solvikenbatforening.portal.styreweb.com/secure/";
	var homePageUrl = baseUrl + "Default2.aspx";
	var marinaUrl   = baseUrl + "archive/marina.aspx";
	var framleieUrl = baseUrl + "archive/marinasublet.aspx";
	var medlemmerUrl = baseUrl + "Members.aspx";
	var grupperingUrl = baseUrl + "Group.aspx?";
	var downloadFolder = $@"C:\Users\{Environment.UserName}\Downloads";

	using (var playwright = await Playwright.CreateAsync())
	{
		var browser = await playwright.Chromium.LaunchAsync();
		var page = await browser.NewPageAsync();
		Console.Write("Logger inn på StyreWeb...");
		await page.GotoAsync(loginUrl);
		await page.GetByLabel("Brukernavn").FillAsync(brukerNavn);
		await page.GetByLabel("Passord").FillAsync(DecodeString(kodetPassord));
		await page.GetByRole(AriaRole.Button).ClickAsync();
		await page.WaitForURLAsync(homePageUrl);
		Console.WriteLine("Logget inn");

		// Do the job
		await DownloadMarina(page, marinaUrl, downloadFolder);
		await DownloadFramleie(page, framleieUrl, downloadFolder);
		await DownloadMedlemmer(page, medlemmerUrl, downloadFolder);
		await DownloadGruppering(page, grupperingUrl, downloadFolder, "Venteliste");
		await DownloadGruppering(page, grupperingUrl, downloadFolder, "Innskudd uten båt");

		//await EndreBatplassStorrelser(page, marinaUrl);
		//await EndreBatplassGrupper(page, marinaUrl);

		//string[] ledigePlasser = {"4V02","4V13","4V18","4V19","4V20","4V21","4V22","4V23","4V27","4V45","4V47","4V48","4V49","4V53","4V56","4V60","4V72","6H09"};
		//foreach (var plass in ledigePlasser)
		//{
		//	await LeverInnBatplass(page, marinaUrl, plass);
		//}
		
		//foreach (var plass in new string[] {"2V09","2H09","4V03","4V04","4V08","4V14","4V29","4V50","4V51","4V52","4V55","4V59","4V61"})
		//{
		//	await EndrePlassVerdier(page, marinaUrl, plass, new List<(string,string)> {("Type", "Ungdomsplass")});
		//}

		//List<(string, string)> ungdomsListe = new List<(string, string)>
		//{
		//	("2V13", "Håkon Skatvedt"),
		//	("5H02", "Marius Noss Gundersen"),
		//	("5V02", "Erik Fadnes Gregersen"),
		//};
		//
		//foreach (var plass in ungdomsListe)
		//{
		//	await FramLeieTilEier(page, marinaUrl, plass.Item1, plass.Item2);
		//}

		//await SettAndelsplasser(page, marinaUrl);
		
		//await ByttPlassSide(page, marinaUrl, downloadFolder);

		await page.ScreenshotAsync(new PageScreenshotOptions { Path = @"C:\MyLocal\Solviken\screenshot.png" });
		await browser.DisposeAsync();
	}
}

static async Task ByttPlassSide(IPage page, string marinaUrl, string downloadFolder)
{
	var marinaFile = Path.Combine(downloadFolder, "Marina.csv");
	using (var reader = new StreamReader(marinaFile))
	{
		string line;
		while ((line = reader.ReadLine()) != null)
		{
			var fields = line.Split('\t');
			if (fields.Length > 2)
			{
				var plass = fields[1];
				if (plass.Length >= 4)
				{
					var plassSide = plass[1];
					string nySide;
					if (plassSide == 'H')
					{
						nySide = "Venstre";
					}
					else if (plassSide == 'V')
					{
						nySide = "Høyre";
					}
					else
					{
						continue;
					}
					
					await EndrePlassVerdier(page, marinaUrl, plass, new List<(string, string)> { ("Side av seksjon", nySide) });
				}
			}
		}
	}
}

static async Task SettAndelsplasser(IPage page, string marinaUrl)
{
	var statusFil = @"C:\Users\Solviken\OneDrive\Solviken\2025\Havnedatabasen\Status før 2026-sesongen.txt";
	using (var reader = new StreamReader(statusFil))
	{
		string line;
		reader.ReadLine();
		reader.ReadLine();
		line = reader.ReadLine();
		int count = int.Parse(line.Split(' ')[0]);
		for (int i = 0; i < count; i++)
		{
			line = reader.ReadLine();
			var plass = line.Split(':')[0];
			await EndrePlassVerdier(page, marinaUrl, plass, new List<(string, string)> {("Type", "Andelsplass")});
		}
	}
}

static async Task EndreBatplassStorrelser(IPage page, string marinaUrl)
{
	var workFolder = @"C:\MyLocal\Solviken";
	var oppmalingPath = @"C:\Users\Solviken\OneDrive\Solviken\2025\Havnedatabasen\Oppmåling-2025";
	var oppmalinger = Directory.GetFiles(oppmalingPath, "*.xlsx");
	foreach (var oppmaling in oppmalinger)
	{
		var csvFile = Path.Combine(workFolder, Path.GetFileNameWithoutExtension(oppmaling) + ".csv");
		ConvertFromXlsx2Csv(oppmaling, csvFile);
		
		using (var reader = new StreamReader(csvFile, Encoding.GetEncoding("UTF-8")))
		{
			reader.ReadLine();		// Skip header
			string line;
			while ((line = reader.ReadLine()) != null)
			{
				var fields = line.Trim().Split('\t');
				if (fields.Length == 4)
				{
					var plassId = fields[0];
					var bredde = GetIntValue(fields[1]) / 100.0;
					var lenH = GetIntValue(fields[2]) / 100;
					var lenV = GetIntValue(fields[3]) / 100;
					var lengde = Math.Max(lenH, lenV);

					await EndrePlassVerdier(page, 
											marinaUrl, 
											plassId, 
											new List<(string, string)> 
											{ 
												("Bredde", bredde.ToString()), 
												("Lengde", lengde.ToString()), 
												("Dybde", "0") 
											});

					//Console.WriteLine($"Plass {plassId}: B={bredde,-8} L={lengde}");
				}
			}
		}
		//await EndrePlassVerdier(page, marinaUrl, "5V16", new List<(string, string)> { ("Bredde", "4,10"), ("Lengde", "10"), ("Dybde", "0") });
	}
}

static async Task EndreBatplassGrupper(IPage page, string marinaUrl)
{
	var hwExportPath = @"C:\Users\Solviken\OneDrive\Solviken\2025\Havnedatabasen\SolvikenBtforening_311224_124226.csv";

	using (var reader = new StreamReader(hwExportPath, Encoding.GetEncoding("UTF-8")))
	{
		reader.ReadLine();      // Skip header
		string line;
		while ((line = reader.ReadLine()) != null)
		{
			var fields = line.Trim().Split('\t');
			if (fields.Length > 10)
			{
				var plassId = fields[2].Substring(0, 4);
				var gruppe = fields[3];
				if (Regex.IsMatch(plassId, @"^[2-6][VH]\d{2}$") &&
					Regex.IsMatch(gruppe, "^[A-L]"))
				{
					await EndrePlassVerdier(page,
											marinaUrl,
											plassId,
											new List<(string, string)>
											{
												("Sted", GetGruppeText(gruppe[0])),
											});
				}


			}
		}
	}
}

static string GetGruppeText(char gruppe)
{
	switch (gruppe)
	{
		case 'A':
			return "Gruppe A: inntil 5,4 m";
		case 'B':
			return "Gruppe B: 5,5 - 7,0 m";
		case 'C':
			return "Gruppe C: 7,1 – 8,7 m";
		case 'D':
			return "Gruppe D: 8,8 – 9,1 m";
		case 'E':
			return "Gruppe E: 9,2 – 10,0 m";
		case 'F':
			return "Gruppe F: 10,1 – 10,6 m";
		case 'G':
			return "Gruppe G: 10,7 – 11,8 m";
		case 'H':
			return "Gruppe H: 11,9 – 12,4 m";
		case 'L':
			return "Gruppe L: 12,5 – 13,7 m";
		default:
			return "";
	}
}

static int GetIntValue(string field)
{
	if (field.Length == 0)
	{
		return 0;
	}
	
	var parts = field.Split('.', ',');
	return int.Parse(parts[0]);
}

static void ConvertFromXlsx2Csv(string excelFile, string csvFile)
{
	if (File.Exists(csvFile))
	{
		if (File.GetLastWriteTime(excelFile) > File.GetLastWriteTime(csvFile))
		{
			File.Delete(csvFile);
		}
		else
		{
			return;
		}
	}

	Console.Write($"Konverterer excel fil {excelFile} til csv fil {csvFile} ...");
	
	string scriptName = @"C:\MyLocal\Solviken\xlsx2csv.vbs"; // full path to script
	ProcessStartInfo ps = new ProcessStartInfo();
	ps.FileName = "cscript.exe";
	ps.Arguments = $"{scriptName} {excelFile} {csvFile}";
	ps.WindowStyle = ProcessWindowStyle.Hidden;
	ps.CreateNoWindow = true;
	var process = Process.Start(ps);
	process.WaitForExit();
	process.Close();
	Console.WriteLine("Ferdig");
}

static async Task DownloadGruppering(IPage page, string grupperingUrl, string downloadFolder, string gruppering)
{
	await page.GotoAsync(grupperingUrl);
	await page.Locator("a").Locator($"text=\"{gruppering}\"").ClickAsync();		// Exact match
	var downloadTask = page.WaitForDownloadAsync();
	await page.ClickAsync("button:has-text(\"Eksport\")");
	var download = await downloadTask;
	var saveFile = gruppering.Replace(' ', '_');
	var savePath = Path.Combine(downloadFolder, $"Gruppe{saveFile}.xlsx");
	await download.SaveAsAsync(savePath);
	Console.WriteLine($"Lastet ned {savePath}");
}

static async Task EndrePlassVerdier(IPage page, string marinaUrl, string plass, List<(string, string)> verdier)
{
	await page.GotoAsync(marinaUrl);
	var inputField = page.Locator("#LeftNavBar_txtSerieNr");
	await inputField.FillAsync(plass);
	await page.ClickAsync("button:has-text(\"Søk\")");

	var tableLocator = page.Locator("sw-panel#pnlMain table#Main_grdv");
	try
	{
		await tableLocator.WaitForAsync(new LocatorWaitForOptions { State = WaitForSelectorState.Visible, Timeout = 5000 });
	}
	catch (Exception)
	{
		Console.WriteLine($"Fant ikke båtplass {plass}");
		return;
	}

	//Console.WriteLine("Table with ID 'Main_grdv' exists inside <sw-panel>.");
	//await page.Locator("a", new PageLocatorOptions { HasTextString = plass }).ClickAsync();
	var plassLenke = page.Locator("a").Locator($"text=\"{plass}\"");
	int hits = await plassLenke.CountAsync();
	if (hits == 0)
	{
		Console.WriteLine($"Fant ikke båtplass {plass} i StyreWeb, hopper over");
		return;
	}
	
	await plassLenke.ClickAsync();
	
	var endreButton = page.Locator("input[type='button'][value='Endre']");
	await endreButton.ClickAsync();
	
	foreach (var verdi in verdier)
	{
		var verdiId = GetEditFieldId(verdi.Item1);
		if (verdiId != null)
		{
			await page.Locator($"#{verdiId}").FillAsync(verdi.Item2);
			Console.WriteLine($"{plass}: {verdi.Item1} = {verdi.Item2}");
		}
		else
		{
			var select = page.Locator($"select[title='{verdi.Item1}']");
			if (select != null)
			{
				await select.SelectOptionAsync(new SelectOptionValue { Label = verdi.Item2 });
				Console.WriteLine($"{plass}: {verdi.Item1} = {verdi.Item2}");
			}
			else
			{
				Console.WriteLine($"{plass}: Fant ikke felt {verdi.Item1}");
				continue;
			}
		}
	}

	var lagreButton = page.Locator("input[type='submit'][value='Lagre']");
	await lagreButton.ClickAsync();
}

static string GetEditFieldId(string fieldName)
{
	switch (fieldName)
	{
		case "Bredde":
			return "Main_details_txtUserDefFlt1";
		case "Lengde":
			return "Main_details_txtUserDefFlt2";
		case "Dybde":
			return "Main_details_txtUserDefFlt3";
		case "Høyde":
			return "Main_details_txtUserDefFlt4";
		case "Innskudd":
			return "Main_details_txtPrice";
	}

	return null;
}

static async Task FramLeieTilEier(IPage page, string marinaUrl, string plass, string eier)
{
	await page.GotoAsync(marinaUrl);
	var inputField = page.Locator("#LeftNavBar_txtSerieNr");
	await inputField.FillAsync(plass);
	await page.ClickAsync("button:has-text(\"Søk\")");

	var tableLocator = page.Locator("sw-panel#pnlMain table#Main_grdv");
	try
	{
		await tableLocator.WaitForAsync(new LocatorWaitForOptions { State = WaitForSelectorState.Visible, Timeout = 5000 });
	}
	catch (Exception)
	{
		Console.WriteLine($"Fant ikke båtplass {plass}");
		return;
	}

	//Console.WriteLine("Table with ID 'Main_grdv' exists inside <sw-panel>.");
	//await page.Locator("a", new PageLocatorOptions { HasTextString = plass }).ClickAsync();
	var plassLenke = page.Locator("a").Locator($"text=\"{plass}\"");
	int hits = await plassLenke.CountAsync();
	if (hits == 0)
	{
		Console.WriteLine($"Fant ikke båtplass {plass} i StyreWeb, hopper over");
		return;
	}

	await plassLenke.ClickAsync();
	
	await page
		.Locator("#Main_grdvMembersSublet tbody tr:first-child a")
		.First
		.ClickAsync();

	// Klikk på endre
	var endreButton = page.Locator("input[type='button'][value='Endre']");
	await endreButton.ClickAsync();

	// Sett Til dato 29.01.2026
	var tilDato = page.Locator("#Main_detailGenericArchiveMember_txtEndDate");
	await tilDato.FillAsync("29.01.2026");

	// Klikk på lagre
	var lagreButton = page.Locator("input[type='submit'][value='Lagre']");
	await lagreButton.ClickAsync();
	Console.WriteLine($"{plass}: Avsluttet framleie fra {eier}");

	// Klikk på "<<"
	await page.GetByText("<<").ClickAsync();

	// Klikk på "Lever inn"
	var leverInn = page.Locator("#Main_btnDeliverIn");
	await leverInn.ClickAsync();

	// Sett til data 29.01.2026
	tilDato = page.Locator("#Main_detailGenericArchiveMember_txtEndDate");
	await tilDato.FillAsync("29.01.2026");

	// Klikk på lagre
	lagreButton = page.Locator("input[type='submit'][value='Lagre']");
	await lagreButton.ClickAsync();
	Console.WriteLine($"{plass}: Levert inn, slettet Solviken som eier");

	// Klikk på "<<"
	await page.GetByText("<<").ClickAsync();

	// Klikk på "Lever ut"
	var leverUt = page.Locator("#Main_btnDeliverOut");
	await leverUt.ClickAsync();

	// Sett medlem
	var medlemSelect = page.Locator("#Main_detailGenericArchiveMember_cboAccDebit");
	await medlemSelect.WaitForAsync(new LocatorWaitForOptions { State = WaitForSelectorState.Visible, Timeout = 5000 });
	var medlemOptions = page.Locator("#Main_detailGenericArchiveMember_cboAccDebit option");

	var options = await medlemOptions.AllInnerTextsAsync();
	string label = options.First(o => o.StartsWith(eier));
	await medlemSelect.SelectOptionAsync(label);

	// Sett fra dato
	var fraDato = page.Locator("#Main_detailGenericArchiveMember_txtStartDate");
	await fraDato.FillAsync("30.01.2026");

	// Velg vare variant "> 7,2 m båtlengde"
	var vareVariant = page.Locator("#Main_detailGenericArchiveMember_cboProductVariant");
	await vareVariant.SelectOptionAsync("> 7,2 m båtlengde");

	// Klikk på "Opprett"
	await page.GetByText("Opprett").ClickAsync();
	Console.WriteLine($"{plass}: Satt {eier} som nye eier av plassen\n");
}
		

static async Task DownloadMarina(IPage page, string marinaUrl, string downloadFolder)
{
	await page.GotoAsync(marinaUrl);
	await page.Locator("#cboAction").SelectOptionAsync(new SelectOptionValue { Label = "Marina - Detaljert" });
	await VisRapportOgLastNed(page, Path.Combine(downloadFolder, "Marina_-_Detaljert.csv"));
}

static async Task DownloadFramleie(IPage page, string framleieUrl, string downloadFolder)
{
	await page.GotoAsync(framleieUrl);
	await VisRapportOgLastNed(page, Path.Combine(downloadFolder, "Fremleie_historie.csv"));
}

static async Task DownloadMedlemmer(IPage page, string medlemmerUrl, string downloadFolder)
{
	await page.GotoAsync(medlemmerUrl);
	await page.Locator("#Main_cboAction").SelectOptionAsync(new SelectOptionValue { Label = "Detaljert Rapport" });
	await VisRapportOgLastNed(page, Path.Combine(downloadFolder, "Detaljert_Rapport.csv"));
}

static async Task LagOpplagsplass(IPage page, string marinaUrl, string felt, int nummer)
{
	await page.GotoAsync(marinaUrl);
	await page.ClickAsync("button:has-text(\"Lag ny\")");

	// Set inn seksjon
	var seksjon = $"Land {felt}";
	var dropdown = page.Locator("select[title='Seksjon']");
	await dropdown.SelectOptionAsync(new SelectOptionValue { Label = seksjon });

	var inputField = page.Locator("#Main_details_txtSortOrder");
	await inputField.FillAsync(nummer.ToString());

	inputField = page.Locator("#Main_details_txtGenericArchiveShortName");
	await inputField.FillAsync($"{felt}{nummer.ToString("d2")}");

	dropdown = page.Locator("select[title='Type']");
	await dropdown.SelectOptionAsync(new SelectOptionValue { Label = "Landopplag" });

	inputField = page.Locator("#Main_details_txtUserDefFlt1");
	await inputField.FillAsync("5");

	inputField = page.Locator("#Main_details_txtUserDefFlt2");
	await inputField.FillAsync("10");

	dropdown = page.Locator("select[title='Vare']");
	await dropdown.SelectOptionAsync(new SelectOptionValue { Label = "Båtplass Avgift" });

	var checkbox = page.Locator("#Main_details_ctl23");
	await checkbox.CheckAsync();

	checkbox = page.Locator("#Main_details_ctl24");
	await checkbox.CheckAsync();

	await page.GetByText("Opprett").ClickAsync();
	
	await page.GetByText("<< Marina").ClickAsync();

	Console.WriteLine($"Opprettet opplagsplass {felt}{nummer.ToString("d2")} på felt {seksjon}");
}

static async Task LeverInnBatplass(IPage page, string marinaUrl, string plass)
{
	await page.GotoAsync(marinaUrl);
	var inputField = page.Locator("#LeftNavBar_txtSerieNr");
	await inputField.FillAsync(plass);
	await page.ClickAsync("button:has-text(\"Søk\")");

	var tableLocator = page.Locator("sw-panel#pnlMain table#Main_grdv");
	try
	{
		await tableLocator.WaitForAsync(new LocatorWaitForOptions { State = WaitForSelectorState.Visible, Timeout = 5000 });
	}
	catch (Exception)
	{
		Console.WriteLine($"Fant ikke båtplass {plass}");
		return;
	}

	//Console.WriteLine("Table with ID 'Main_grdv' exists inside <sw-panel>.");
	//await page.Locator("a", new PageLocatorOptions { HasTextString = plass }).ClickAsync();
	var plassLenke = page.Locator("a").Locator($"text=\"{plass}\"");
	int hits = await plassLenke.CountAsync();
	if (hits == 0)
	{
		Console.WriteLine($"Fant ikke båtplass {plass} i StyreWeb, hopper over");
		return;
	}

	await plassLenke.ClickAsync();
	var leverInn = page.Locator("#Main_btnDeliverIn");
	await leverInn.ClickAsync();
	var tilDato = page.Locator("#Main_detailGenericArchiveMember_txtEndDate");
	var today = DateTime.Now.ToString("d.MM.yyyy");
	await tilDato.FillAsync(today);
	
	var lagreButton = page.Locator("input[type='submit'][value='Lagre']");
	await lagreButton.ClickAsync();
	Console.WriteLine($"Levert inn {plass}");
}

static async Task VisRapportOgLastNed(IPage page, string savePath)
{
	Console.Write("Genererer rapport...");
	var newPageTask = page.Context.WaitForPageAsync();
	await page.ClickAsync("button:has-text(\"Vis\")");
	var newTab = await newPageTask;
	await newTab.WaitForLoadStateAsync(LoadState.DOMContentLoaded);
	Console.Write("Ferdig...");

	Console.Write("Laster ned...");
	var downloadTask = newTab.WaitForDownloadAsync();
	await newTab.ClickAsync("button:has-text(\"Eksporter som csv fil\")");
	var download = await downloadTask;
	await download.SaveAsAsync(savePath);
	Console.WriteLine(savePath);
	await newTab.CloseAsync();
}

static string DecodeString(string coded)
{
	string alfaNum = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789abcdefghijklmnopqrstuvwxyz";
	int maxPos = alfaNum.Length - 1;
	var decoded = new List<char>();
	
	foreach (var c in coded)
	{
		int index = alfaNum.IndexOf(c);
		int newIndex = index - 52;
		if (newIndex < 0)
		{
			newIndex += maxPos;
		}
		newIndex = newIndex % maxPos;
		decoded.Add(alfaNum[newIndex]);
	}
	
	return new string(decoded.ToArray());
}