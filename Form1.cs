namespace LatinScrapper
{
    using PuppeteerSharp;
    public class LatinData
    {
        public string Word;
        public string Meaning;
    }

    public partial class Form1 : Form
    {
        private readonly string siteUrl = "https://www.dl.cambridgescp.com/Array/book-ii-stage-stage-teachers-guide";
        private readonly string exportPath;
        private Browser _browser;
        private Page _page;
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
        // TODO get conjuration and meaning of the current word
        //var getWords = @"const words = [...document.querySelectorAll('span')].filter(e => {
        //      return e.id.match(/w\d+/gm)
        //})";

        //let word = document.querySelector('span#w0.w').innerText
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
                    await page.ClickAsync($"span#anay3.loli");
                    await page.ClickAsync($"span#w{wordIndex}");
                    while (wordIndex < wordsCount)
                    {
                        var word = await page.WaitForSelectorAsync($"span#w{wordIndex++}");
                        var wordText = (string)await word.EvaluateFunctionAsync("el => el.textContent");
                        var meaning = await page.WaitForSelectorAsync($"div#prs");
                        var meaningText = (string)await meaning.EvaluateFunctionAsync("el => el.textContent");
                        data.Add(new LatinData
                        {
                            Word = wordText,
                            Meaning = meaningText,
                        });
                        var keyboard = page.Keyboard;
                        await keyboard.PressAsync("ArrowRight");
                    }
                }
                finally
                {
                    await _browser.CloseAsync();
                    _browser?.Dispose();
                }
                
            }
            
        }
    }
}