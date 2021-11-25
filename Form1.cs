namespace LatinScrapper
{
    using Newtonsoft.Json;
    using OfficeOpenXml;
    using PuppeteerSharp;

    public partial class Form1 : Form
    {
        private readonly string siteUrl = "https://www.dl.cambridgescp.com/Array/book-ii-stage-stage-teachers-guide";
        private Browser? _browser;
        private Page _page;
        public Form1()
        {
            InitializeComponent();
        }

        private async void button1_Click(object sender, EventArgs e)
        {
            if (_browser == null)
            {
                _browser = await Puppeteer.LaunchAsync(new LaunchOptions
                {
                    Headless = false,
                    ExecutablePath = $"C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
                });
                _page = await _browser.NewPageAsync();
                await _page.SetViewportAsync(new ViewPortOptions { Width = 1200, Height = 800 });
                await _page.GoToAsync(siteUrl);
            }
           else
            {
                MessageBox.Show("Browser is already opened", "Error");
            }
           
        }
        
        private async void button2_Click(object sender, EventArgs e)
        {
            if(_browser == null)
            {
                MessageBox.Show("Open browser first", "Error");
                return;
            }
            if (filePath.Text == "")
            {
                MessageBox.Show("Select output path first", "Error");
                return;

            }
            
            var pages = await _browser.PagesAsync();
            List<LatinData> data = new();
            if (pages.Length > 1)
            {
                var page = pages.FirstOrDefault(x => x.Url.Contains("https://www.dl.cambridgescp.com/sites"));
                if (page == null)
                {
                    MessageBox.Show("Open browser page to scrap", "Error");
                    return;

                }
                var wordIndex = 0;
                try
                {
                    var documentNameNode = await page.WaitForSelectorAsync($"div#hdr");
                    var documentName = (string)await documentNameNode.EvaluateFunctionAsync(@"el => {
                        if (el.firstElementChild) {
                            return el.firstElementChild.previousSibling.nodeValue.trim();
                        } else {
                            return el.textContent;
                        }
                    }");
                    var wordsCount = (int)await page.EvaluateFunctionAsync(@"() => [...document.querySelectorAll('span')].filter(e => e.id.match(/w\d+/gm)).length");
                    await page.ClickAsync($"span#w{wordIndex}");
                    while (wordIndex < wordsCount)
                    {
                        var word = await page.WaitForSelectorAsync($"span#w{wordIndex++}");
                        var wordText = (string)await word.EvaluateFunctionAsync("el => el.textContent");
                        var meaning = await page.WaitForSelectorAsync($"div#prs");
                        var scrappedData = await meaning.EvaluateFunctionAsync(@"el =>{
                                let res
                                switch (el.children.length) {
                                    case 1 : {
                                        let c = el.firstElementChild.textContent;
                                        let grammar =  el.firstElementChild.nextSibling.nodeValue.trim();
                                        res = [{Conjugation: c, Meaning: grammar}]
                                        break;
                                    }
                                    case 2 : {
                                        let c = el.firstElementChild.textContent;
                                        let grammar =  el.firstElementChild.nextSibling.nodeValue.trim();
                                        let g = el.lastElementChild.textContent;
                                        res = [{Conjugation: c, Meaning: grammar, Analysis: g}]
                                        break;
                                    }
                                    case 3 : {
                                       let w1 = el.firstElementChild.textContent;
                                       let m1 = el.firstElementChild.nextSibling.nodeValue.trim();
                                       let w2 = el.children[2].textContent;
                                       let m2 = el.children[2].nextSibling.nodeValue.trim();
                                       res = [{Conjugation: w1, Meaning: m1},
                                              {Word: w2, Meaning: m2}]
                                        break;
                                    }
                                    case 4 : {
                                       let w1 = el.firstElementChild.textContent;
                                       let m1 = el.firstElementChild.nextSibling.nodeValue.trim();
                                       let w2 = el.children[2].textContent;
                                       let m2 = el.children[2].nextSibling.nodeValue.trim();
                                       let analys = el.lastElementChild.textContent;
                                       res = [{Conjugation: w1, Meaning: m1, Analysis: analys},
                                              {Word: w2, Meaning: m2}]
                                       break;
                                    }
                                    default : {
                                        res = [{Conjugation: el.textContent}];
                                    }
                                }
                                return res;
                            }");
                        var scrappedDataString = JsonConvert.SerializeObject(scrappedData);
                        var result = JsonConvert.DeserializeObject<List<LatinData>>(scrappedDataString);
                        result.ForEach(x => { 
                            if (x.Word == null)
                            {
                                x.Word = wordText;
                            }
                        });
                        data.AddRange(result);
                        var keyboard = page.Keyboard;
                        await keyboard.PressAsync("ArrowRight");
                    }
                    var distinctData = data.DistinctBy(x => x.Word).OrderBy(x => x.Word);
                    ExportToExcel(distinctData, $"{filePath.Text}\\{documentName}.xlsx");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    await _browser.CloseAsync();
                    _browser?.Dispose();
                    _browser = null;
                }

            }
            
        }

        private void ExportToExcel(IEnumerable<LatinData> latinData, string path)
        {
            var outputStream = new MemoryStream();
            try
            {
                using(var package = new ExcelPackage(outputStream))
                {
                    var ws = package.Workbook.Worksheets.Add("Vocabulary");
                    var dataRange = ws.Cells["A1"].LoadFromCollection(latinData, true, OfficeOpenXml.Table.TableStyles.Medium1);
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

        private void button3_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            fbd.RootFolder = Environment.SpecialFolder.MyDocuments;
            fbd.Description = "Select folder";
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                filePath.Text = fbd.SelectedPath;
            }
        }

        private async void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if(_browser != null)
            {
                await _browser.CloseAsync();
                _browser?.Dispose();
                _browser = null;
            }
        }
    }
    public class LatinData
    {
        public string? Word { get; set; }
        public string? Conjugation { get; set; }
        public string? Meaning { get; set; }
        public string? Analysis { get; set; }
    }
}