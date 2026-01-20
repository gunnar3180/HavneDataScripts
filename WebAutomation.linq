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

		//await EndrePlassVerdier(page, marinaUrl, "5V16", new List<(string, string)> {("Innskudd", "16450"), ("Dybde", "1,90")});
		
		await page.ScreenshotAsync(new PageScreenshotOptions { Path = @"C:\MyLocal\Solviken\screenshot.png" });
		await browser.DisposeAsync();
	}
}

static async Task DownloadGruppering(IPage page, string grupperingUrl, string downloadFolder, string gruppering)
{
	await page.GotoAsync(grupperingUrl);
	await page.Locator("a").Locator($"text=\"{gruppering}\"").ClickAsync();		// Exact match
	var downloadTask = page.WaitForDownloadAsync();
	await page.ClickAsync("button:has-text(\"Eksport\")");
	var download = await downloadTask;
	var savePath = Path.Combine(downloadFolder, $"Gruppe{gruppering}.xlsx");
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
	await page.Locator("a", new PageLocatorOptions { HasTextString = plass }).ClickAsync();
	
	var endreButton = page.Locator("input[type='button'][value='Endre']");
	await endreButton.ClickAsync();
	
	foreach (var verdi in verdier)
	{
		var verdiId = GetEditFieldId(verdi.Item1);
		if (verdiId == null)
		{
			Console.WriteLine($"{plass}: Fant ikke felt {verdi.Item1}");
			continue;
		}

		await page.Locator($"#{verdiId}").FillAsync(verdi.Item2);
		Console.WriteLine($"{plass}: {verdi.Item1} = {verdi.Item2}");
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