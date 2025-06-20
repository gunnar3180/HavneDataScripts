<Query Kind="Program">
  <NuGetReference>Microsoft.Playwright</NuGetReference>
  <Namespace>System.Threading.Tasks</Namespace>
  <Namespace>Microsoft.Playwright</Namespace>
</Query>

static async Task Main()
{
	Microsoft.Playwright.Program.Main(new[] { "install" });   // Install Playwright if not already installed
	await Go().ConfigureAwait(false);
}

public static async Task Go()
{
	using (var playwright = await Playwright.CreateAsync())
	{
		var browser = await playwright.Chromium.LaunchAsync();
		var page = await browser.NewPageAsync();
		await page.GotoAsync("https://playwright.dev/dotnet");
		await page.ScreenshotAsync(new PageScreenshotOptions { Path = @"C:\MyLocal\Solviken\screenshot.png" });
		await browser.DisposeAsync();
	}
}

