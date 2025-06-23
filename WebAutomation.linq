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
	var hjemUrl = "https://solvikenbatforening.portal.styreweb.com/secure/";
	var marinaUrl   = hjemUrl + "archive/marina.aspx";
	var framleieUrl = hjemUrl + "archive/marinasublet.aspx";
	var medlemmerUrl = hjemUrl + "Members.aspx";
	var downloadFolder = @"C:\Users\solvi\Downloads";

	using (var playwright = await Playwright.CreateAsync())
	{
		var browser = await playwright.Chromium.LaunchAsync();
		var page = await browser.NewPageAsync();
		await page.GotoAsync(loginUrl);
		await page.GetByLabel("Brukernavn").FillAsync(brukerNavn);
		await page.GetByLabel("Passord").FillAsync(DecodeString(kodetPassord));
		await page.GetByRole(AriaRole.Button).ClickAsync();
		await page.WaitForURLAsync(hjemUrl);
		Console.WriteLine("Logget inn på StyreWeb");
		
		await page.GotoAsync(marinaUrl);
		await VisRapportOgLastNed(page, Path.Combine(downloadFolder, "Marina.csv"));

		await page.Locator("#cboAction").SelectOptionAsync(new SelectOptionValue { Label = "Marina - Detaljert" });
		await VisRapportOgLastNed(page, Path.Combine(downloadFolder, "Marina_-_Detaljert.csv"));

		await page.GotoAsync(framleieUrl);
		await VisRapportOgLastNed(page, Path.Combine(downloadFolder, "Fremleie_historie.csv"));

		await page.GotoAsync(medlemmerUrl);
		await page.Locator("#Main_cboAction").SelectOptionAsync(new SelectOptionValue { Label = "Detaljert Rapport" });
		await VisRapportOgLastNed(page, Path.Combine(downloadFolder, "Detaljert_Rapport.csv"));

		//await page.ScreenshotAsync(new PageScreenshotOptions { Path = @"C:\MyLocal\Solviken\screenshot.png" });
		await browser.DisposeAsync();
	}
}

static async Task VisRapportOgLastNed(IPage page, string savePath)
{
	Console.Write("Genererer rapport...");
	var newPageTask = page.Context.WaitForPageAsync();
	await page.ClickAsync("button:has-text(\"Vis\")");
	var newTab = await newPageTask;
	await newTab.WaitForLoadStateAsync(LoadState.DOMContentLoaded);
	Console.WriteLine("Ferdig");

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