namespace LatinScrapper
{
    using OfficeOpenXml;
    using PuppeteerSharp;
    public class LatinData
    {
        public string Word { get; set; }
        public string Conjuration { get; set; }
        public string Meaning { get; set; }
        public string Analysis { get; set; }
    }

    public partial class Form1 : Form
    {
        private readonly string siteUrl = "https://www.dl.cambridgescp.com/Array/book-ii-stage-stage-teachers-guide";
        private readonly string exportPath;
        private Browser _browser;
        private PuppeteerSharp.Page _page;
        public Form1()
        {
            InitializeComponent();
        }

        private async void button1_Click(object sender, EventArgs e)
        {
            _browser = await Puppeteer.LaunchAsync(new LaunchOptions
            {
                Headless = false,
                ExecutablePath = $"C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
            });
            _page = await _browser.NewPageAsync();
            await _page.GoToAsync(siteUrl);
           
        }
 
        private async void button2_Click(object sender, EventArgs e)
        {
            var pages = await _browser.PagesAsync();
            List<LatinData> data = new List<LatinData>();
            if (pages.Length > 1)
            {
                var page = pages.First(x => x.Url.Contains("https://www.dl.cambridgescp.com/sites"));
                var wordIndex = 0;
                try
                {
                    var wordsCount = (int)await page.EvaluateFunctionAsync(@"() => [...document.querySelectorAll('span')].filter(e => e.id.match(/w\d+/gm)).length");
                    await page.ClickAsync($"span#w{wordIndex}");
                    while (wordIndex < wordsCount)
                    {
                        var word = await page.WaitForSelectorAsync($"span#w{wordIndex++}");
                        var wordText = (string)await word.EvaluateFunctionAsync("el => el.textContent");
                        var meaning = await page.WaitForSelectorAsync($"div#prs");
                        var conjText = (string)await meaning.EvaluateFunctionAsync("el => el.firstElementChild.textContent");
                        var analysisText = (string)await meaning.EvaluateFunctionAsync("el => el.lastElementChild.textContent");
                        var meaningText = (string)await meaning.EvaluateFunctionAsync(@"el =>{
                            const c = el.firstElementChild.textContent;
                            const g = el.lastElementChild.textContent;
                            return el.textContent.split(c)[1].split(g)[0].trim();
                            }");

                        data.Add(new LatinData
                        {
                            Word = wordText,
                            Conjuration = conjText,
                            Meaning = meaningText,
                            Analysis = analysisText,
                        });
                        var keyboard = page.Keyboard;
                        await keyboard.PressAsync("ArrowRight");
                    }
                    ExportToExcel(data, @"C:\projects\test1.xlsx");

                }
                finally
                {
                    await _browser.CloseAsync();
                    _browser?.Dispose();
                }
                
            }
            
        }

        private void ExportToExcel(List<LatinData> latinData, string path)
        {
            var outputStream = new MemoryStream();
            try
            {
                using(var package = new ExcelPackage(outputStream))
                {
                    var ws = package.Workbook.Worksheets.Add("Vocabulary");
                    var dataRange = ws.Cells["A1"].LoadFromCollection(latinData, true);
                    ws.Column(1).AutoFit();
                    ws.Column(2).AutoFit();
                    ws.Column(3).AutoFit();
                    ws.Column(4).AutoFit();
                    package.Save();
                    outputStream.Position = 0;

                    byte[] data = package.GetAsByteArray();
                    File.WriteAllBytes(path, data);
                }
            }
            catch
            {
                outputStream.Dispose();
                return;
            }
        }
    }
}