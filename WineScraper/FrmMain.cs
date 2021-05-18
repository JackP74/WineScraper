using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;

using HtmlAgilityPack;
using MessageCustomHandler;
using Ookii.Dialogs.WinForms;
using HtmlDocument = HtmlAgilityPack.HtmlDocument;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text;

namespace WineScraper
{
    public partial class FrmMain : Form
    {
        #region "Variables"
        private readonly string AppPath = Application.StartupPath;
        private readonly string WineUrl = @"https://www.maccaninodrink.com";
        private Thread scraperThread;
        private readonly string newLine = Environment.NewLine;
        #endregion

        #region "Win32 Imports"
        [DllImport("kernel32.dll")]
        public static extern Int32 AllocConsole();

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern IntPtr GetStdHandle(int nStdHandle);

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern bool SetStdHandle(int nStdHandle, IntPtr hHandle);

        public const int STD_OUTPUT_HANDLE = -11;
        public const int STD_INPUT_HANDLE = -10;
        public const int STD_ERROR_HANDLE = -12;

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern IntPtr CreateFile([MarshalAs(UnmanagedType.LPTStr)] string filename,
                                               [MarshalAs(UnmanagedType.U4)] uint access,
                                               [MarshalAs(UnmanagedType.U4)] FileShare share,
                                                                                 IntPtr securityAttributes,
                                               [MarshalAs(UnmanagedType.U4)] FileMode creationDisposition,
                                               [MarshalAs(UnmanagedType.U4)] FileAttributes flagsAndAttributes,
                                                                                 IntPtr templateFile);

        public const uint GENERIC_WRITE = 0x40000000;
        public const uint GENERIC_READ = 0x80000000;

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern Int32 SendMessage(IntPtr hWnd, int msg, int wParam, [MarshalAs(UnmanagedType.LPWStr)] string lParam);

        private const int EM_SETCUEBANNER = 0x1501;

        private static void OverrideRedirection()
        {
            var hOut = GetStdHandle(STD_OUTPUT_HANDLE);
            var hRealOut = CreateFile("CONOUT$", GENERIC_READ | GENERIC_WRITE, FileShare.Write, IntPtr.Zero, FileMode.OpenOrCreate, 0, IntPtr.Zero);
            if (hRealOut != hOut)
            {
                SetStdHandle(STD_OUTPUT_HANDLE, hRealOut);
                Console.SetOut(new StreamWriter(Console.OpenStandardOutput(), Console.OutputEncoding) { AutoFlush = true });
            }
        }
        #endregion

        #region "Multi-Threading"
        private delegate void SetControlPropertyThreadSafeDelegate(Control control, string propertyName, object propertyValue);

        private void SetControlProperty(Control control, string propertyName, object propertyValue)
        {
            if (control.InvokeRequired)
                control.Invoke(new SetControlPropertyThreadSafeDelegate(SetControlProperty), new object[] { control, propertyName, propertyValue });
            else
                control.GetType().InvokeMember(propertyName, BindingFlags.SetProperty, null, control, new object[] { propertyValue });
        }
        #endregion

        #region "Functions"
        public FrmMain()
        {
            InitializeComponent();
            CreateConsole();

            SetCueText(TxtSavePath, "Save path...");
        }

        private void StartThread(ThreadStart newStart)
        {
            Thread newThread = new Thread(newStart) { IsBackground = true };
            newThread.SetApartmentState(ApartmentState.STA);
            newThread.Start();
        }

        private void CreateConsole()
        {
            AllocConsole();

            OverrideRedirection();

            Console.Title = "Wine Console";
            Console.ForegroundColor = ConsoleColor.Red;
            Console.Write("Wine Scraper");
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine(" v0.1 alpha");
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine("");
        }

        private void Log(string txt)
        {
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine(txt);
        }

        private void LogWarning(string warning)
        {
            Console.ForegroundColor = ConsoleColor.DarkYellow;
            Console.WriteLine(warning);
            Console.ForegroundColor = ConsoleColor.White;
        }

        private void LogError(string error)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine(error);
            Console.ForegroundColor = ConsoleColor.White;
        }

        private void SetStatus(string newStatus)
        {
            SetControlProperty(LabelStatus, "Text", "Status: " + newStatus);
        }

        private void SetCueText(Control control, string text)
        {
            SendMessage(control.Handle, EM_SETCUEBANNER, 0, text);
        }

        private string MakeValidFileName(string name)
        {
            string invalidChars = Regex.Escape(new string(Path.GetInvalidFileNameChars()));
            string invalidRegStr = string.Format(@"([{0}]*\.+$)|([{0}]+)", invalidChars);

            return Regex.Replace(name, invalidRegStr, "_");
        }

        private void SetUpHeaders(Excel.Worksheet xlWorkSheet)
        {
            xlWorkSheet.Cells[1, 1] = "Handle";
            xlWorkSheet.Cells[1, 2] = "Title";
            xlWorkSheet.Cells[1, 3] = "Body (HTML)";
            xlWorkSheet.Cells[1, 4] = "Vendor";
            xlWorkSheet.Cells[1, 5] = "Type";
            xlWorkSheet.Cells[1, 6] = "Tags";
            xlWorkSheet.Cells[1, 7] = "Published";
            xlWorkSheet.Cells[1, 8] = "Option1 Name";
            xlWorkSheet.Cells[1, 9] = "Option1 Value";
            xlWorkSheet.Cells[1, 10] = "Option2 Name";
            xlWorkSheet.Cells[1, 11] = "Option2 Value";
            xlWorkSheet.Cells[1, 12] = "Option3 Name";
            xlWorkSheet.Cells[1, 13] = "Option3 Value";
            xlWorkSheet.Cells[1, 14] = "Variant SKU";
            xlWorkSheet.Cells[1, 15] = "Variant Grams";
            xlWorkSheet.Cells[1, 16] = "Variant Inventory Tracker";
            xlWorkSheet.Cells[1, 17] = "Variant Inventory Qty";
            xlWorkSheet.Cells[1, 18] = "Variant Inventory Policy";
            xlWorkSheet.Cells[1, 19] = "Variant Fulfillment Service";
            xlWorkSheet.Cells[1, 20] = "Variant Price";
            xlWorkSheet.Cells[1, 21] = "Variant Compare At Price";
            xlWorkSheet.Cells[1, 22] = "Variant Requires Shipping";
            xlWorkSheet.Cells[1, 23] = "Variant Taxable";
            xlWorkSheet.Cells[1, 24] = "Variant Barcode";
            xlWorkSheet.Cells[1, 25] = "Image Src";
            xlWorkSheet.Cells[1, 26] = "Image Position";
            xlWorkSheet.Cells[1, 27] = "Image Alt Text";
            xlWorkSheet.Cells[1, 28] = "Gift Card";
            xlWorkSheet.Cells[1, 29] = "SEO Title";
            xlWorkSheet.Cells[1, 30] = "SEO Description";
            xlWorkSheet.Cells[1, 31] = "Google Shopping / Google Product Category";
            xlWorkSheet.Cells[1, 32] = "Google Shopping / Gender";
            xlWorkSheet.Cells[1, 33] = "Google Shopping / Age Group";
            xlWorkSheet.Cells[1, 34] = "Google Shopping / MPN";
            xlWorkSheet.Cells[1, 35] = "Google Shopping / AdWords Grouping";
            xlWorkSheet.Cells[1, 36] = "Google Shopping / AdWords Labels";
            xlWorkSheet.Cells[1, 37] = "Google Shopping / Condition";
            xlWorkSheet.Cells[1, 38] = "Google Shopping / Custom Product";
            xlWorkSheet.Cells[1, 39] = "Google Shopping / Custom Label 0";
            xlWorkSheet.Cells[1, 40] = "Google Shopping / Custom Label 1";
            xlWorkSheet.Cells[1, 41] = "Google Shopping / Custom Label 2";
            xlWorkSheet.Cells[1, 42] = "Google Shopping / Custom Label 3";
            xlWorkSheet.Cells[1, 43] = "Google Shopping / Custom Label 4";
            xlWorkSheet.Cells[1, 44] = "Variant Image";
            xlWorkSheet.Cells[1, 45] = "Variant Weight Unit";
            xlWorkSheet.Cells[1, 46] = "Variant Tax Code";
            xlWorkSheet.Cells[1, 47] = "Cost per item";
            xlWorkSheet.Cells[1, 48] = "Status";
        }

        private void SetProduct(Excel.Worksheet xlWorkSheet, int row, string title, string slug, string description, string type, string tags, string price, string imgURL, string seoTitle, string seoDescription)
        {
            xlWorkSheet.Cells[row, 1] = slug;
            xlWorkSheet.Cells[row, 2] = title;
            xlWorkSheet.Cells[row, 3] = description;
            xlWorkSheet.Cells[row, 4] = "";
            xlWorkSheet.Cells[row, 5] = type;
            xlWorkSheet.Cells[row, 6] = tags;
            xlWorkSheet.Cells[row, 7] = "Published";
            xlWorkSheet.Cells[row, 8] = title;
            xlWorkSheet.Cells[row, 9] = "Default " + title;
            xlWorkSheet.Cells[row, 10] = "";
            xlWorkSheet.Cells[row, 11] = "";
            xlWorkSheet.Cells[row, 12] = "";
            xlWorkSheet.Cells[row, 13] = "";
            xlWorkSheet.Cells[row, 14] = "";
            xlWorkSheet.Cells[row, 15] = "1000";
            xlWorkSheet.Cells[row, 16] = "shopify";
            xlWorkSheet.Cells[row, 17] = "";
            xlWorkSheet.Cells[row, 18] = "continue";
            xlWorkSheet.Cells[row, 19] = "manual";
            xlWorkSheet.Cells[row, 20] = price;
            xlWorkSheet.Cells[row, 21] = price;
            xlWorkSheet.Cells[row, 22] = "TRUE";
            xlWorkSheet.Cells[row, 23] = "FALSE";
            xlWorkSheet.Cells[row, 24] = "";
            xlWorkSheet.Cells[row, 25] = imgURL;
            xlWorkSheet.Cells[row, 26] = "1";
            xlWorkSheet.Cells[row, 27] = "";
            xlWorkSheet.Cells[row, 28] = "FALSE";
            xlWorkSheet.Cells[row, 29] = seoTitle;
            xlWorkSheet.Cells[row, 30] = seoDescription;
            xlWorkSheet.Cells[row, 31] = "";
            xlWorkSheet.Cells[row, 32] = "";
            xlWorkSheet.Cells[row, 33] = "";
            xlWorkSheet.Cells[row, 34] = "";
            xlWorkSheet.Cells[row, 35] = "";
            xlWorkSheet.Cells[row, 36] = "";
            xlWorkSheet.Cells[row, 37] = "";
            xlWorkSheet.Cells[row, 38] = "";
            xlWorkSheet.Cells[row, 39] = "";
            xlWorkSheet.Cells[row, 40] = "";
            xlWorkSheet.Cells[row, 41] = "";
            xlWorkSheet.Cells[row, 42] = "";
            xlWorkSheet.Cells[row, 43] = "";
            xlWorkSheet.Cells[row, 44] = "";
            xlWorkSheet.Cells[row, 45] = "g";
            xlWorkSheet.Cells[row, 46] = "";
            xlWorkSheet.Cells[row, 47] = "";
            xlWorkSheet.Cells[row, 48] = "active";
        }

        private void Scrape()
        {
            try
            {
                Log("Getting source...");
                SetStatus("Working...");

                string MainDir = TxtSavePath.Text;

                if (string.IsNullOrWhiteSpace(MainDir) || !Directory.Exists(MainDir))
                {
                    MainDir = Path.Combine(AppPath, @"Products");

                    if (!Directory.Exists(MainDir))
                        Directory.CreateDirectory(MainDir);
                }

                using WebClient client = new WebClient();
                client.Encoding = System.Text.Encoding.UTF8;
                client.Headers.Add("user-agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:88.0) Gecko/20100101 Firefox/88.0");

                var doc = new HtmlDocument();
                string rawHtml = client.DownloadString(WineUrl);

                if (string.IsNullOrWhiteSpace(rawHtml))
                {
                    LogError("No response from server");
                    return;
                }

                doc.LoadHtml(rawHtml);

                var MenuNode = doc.GetElementbyId("menu_ver_2");
                var LinksNode = MenuNode.SelectSingleNode(".//div[@class='dropdown-menus']/ul");

                var LinkNodes = LinksNode.SelectNodes(".//li");

                List<string> categoryURLs = new List<string>();

                foreach (var Link in LinkNodes)
                {
                    var hrefNode = Link.SelectSingleNode(".//a[@href]");
                    string itemLink = hrefNode.GetAttributeValue("href", string.Empty);

                    categoryURLs.Add($"{itemLink}?limit=2000");
                }

                Excel.Application excel = new Excel.Application
                {
                    Visible = false,
                    DisplayAlerts = false
                };

                if (excel == null)
                {
                    CMBox.Show("Warning", "Excel is not properly installed!", Style.Warning, Buttons.OK);
                    return;
                }

                object misValue = Missing.Value;
                var xlWorkBook = excel.Workbooks.Add(misValue);
                var xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                xlWorkSheet.Name = @"maccaninodrink data";

                SetUpHeaders(xlWorkSheet);

                int CurrentRow = 2;

                foreach (var catURL in categoryURLs)
                {
                    try
                    {
                        Log($"Getting page '{catURL}'");

                        rawHtml = client.DownloadString(catURL);

                        if (string.IsNullOrWhiteSpace(rawHtml))
                        {
                            LogError("No response from server");
                            return;
                        }

                        doc.LoadHtml(rawHtml);

                        Log("Getting product pages...");

                        List<string> productsURLs = new List<string>();

                        var parentElement = doc.GetElementbyId("content");
                        var productsDivs = parentElement.SelectNodes(".//div[contains(@class, 'product-thumb')]");

                        if (productsDivs == null)
                        {
                            Log("Page empty, skipped");
                            continue;
                        }

                        foreach (var product in productsDivs)
                        {
                            string url = product.SelectSingleNode(".//div[@class='name']").SelectSingleNode(".//a[@href]").GetAttributeValue("href", string.Empty);

                            if (url.EndsWith("?limit=100"))
                                url = url.Substring(0, url.Length - 10);

                            bool result = Uri.TryCreate(url, UriKind.Absolute, out Uri uriResult) && (uriResult.Scheme == Uri.UriSchemeHttp || uriResult.Scheme == Uri.UriSchemeHttps);

                            if (result)
                                productsURLs.Add(url);
                        }

                        Log("Getting products...");

                        foreach (var url in productsURLs)
                        {
                            try
                            {
                                rawHtml = client.DownloadString(url);

                                doc = new HtmlDocument();
                                doc.LoadHtml(rawHtml);

                                parentElement = doc.GetElementbyId("content");
                                var productDiv = parentElement.SelectSingleNode(".//div[@class='row']").SelectSingleNode(".//div[@class='product-buy-wrapper']");

                                string name = productDiv.SelectSingleNode(".//h1").InnerText;

                                string validFileName = MakeValidFileName(name);
                                //string productPath = Path.Combine(MainDir, validFileName);

                                // if (!Directory.Exists(productPath))
                                //    Directory.CreateDirectory(productPath);

                                var imgDiv = doc.GetElementbyId("zoom1");
                                string imgUrl = imgDiv.GetAttributeValue("href", string.Empty);
                                imgUrl = imgUrl.Replace("image/cache/catalog", "image/catalog");
                                imgUrl = imgUrl.Replace("image/cache/data", "image/data");
                                imgUrl = imgUrl.Substring(0, imgUrl.Length - 12) + "." + imgUrl.Split('.').Last();

                                //string imgFileName = MakeValidFileName(imgUrl.Split('/').Last());

                                //client.DownloadFile(imgUrl, Path.Combine(productPath, imgFileName));

                                var miscInfoDiv = productDiv.SelectSingleNode(".//div[contains(@class, 'product-buy-logo')]").SelectSingleNode(".//ul[contains(@class, 'list-unstyled')]");

                                var miscDivs = miscInfoDiv.SelectNodes(".//li");

                                string brand = string.Empty;
                                string productCode = string.Empty;

                                foreach (var misc in miscDivs)
                                {
                                    try
                                    {
                                        string innerTxt = misc.InnerText.Trim().Replace("\n", "").Replace("\r", "").Trim().Replace("  ", " ");

                                        if (innerTxt.StartsWith("Brand:"))
                                        {
                                            brand = innerTxt.Substring(6).Trim();
                                        }
                                        else if (innerTxt.StartsWith("Product Code:"))
                                        {
                                            productCode = innerTxt.Substring(13).Trim();
                                        }
                                    }
                                    catch { }
                                }

                                string price = string.Empty;

                                try
                                {
                                    var priceNode = doc.DocumentNode.SelectSingleNode("//div[@class='price-h']");
                                    price = priceNode.InnerText.Trim();
                                }
                                catch { }

                                string exTaxPrice = string.Empty;

                                try
                                {
                                    var exTaxParent = parentElement.SelectSingleNode(".//ul[@class='list-unstyled pp']/li[2]");
                                    exTaxPrice = exTaxParent.InnerText.Trim();
                                    exTaxPrice = exTaxPrice.Split(':')[1].Trim();
                                }
                                catch { }

                                string tags = string.Empty;

                                try
                                {
                                    var tagsParent = parentElement.SelectSingleNode(".//ul[@class='list-unstyled pf pf-bottom']/li[2]");
                                    var tagsNode = tagsParent.SelectNodes(".//a");

                                    List<string> tagList = tagsNode.Where(x => { return !string.IsNullOrWhiteSpace(x.InnerText.Trim()); }).Select(x => x.InnerText.Trim()).ToList();
                                    tags = string.Join(", ", tagList);

                                }
                                catch { }

                                string description = string.Empty;

                                try
                                {
                                    var descNode = doc.GetElementbyId("tab-description");
                                    description = descNode.InnerText.Trim();
                                }
                                catch { }

                                string type = string.Empty;

                                try
                                {
                                    var breadcrumb = doc.DocumentNode.SelectSingleNode(".//ul[@class='breadcrumb']").InnerText.Trim().Replace(" ", "");

                                    RegexOptions options = RegexOptions.None;
                                    Regex regex = new Regex("[ ]{2,}", options);
                                    breadcrumb = regex.Replace(breadcrumb, " ");

                                    List<string> catNames = breadcrumb.Split(new[] { '\r', '\n' }, StringSplitOptions.None).ToList();

                                    catNames.RemoveAt(0);
                                    catNames.RemoveAt(catNames.Count - 1);

                                    string catNamesStr = string.Join(", ", catNames);

                                    tags = $"{catNamesStr}, {tags}";

                                    type = catNames.Last();
                                }
                                catch { }

                                tags = $"\"{tags}\"";

                                string slug = name.ToLower().Replace(" ", "-").RemoveSpecialCharacters();

                                string seoTitle = name.Replace("\n", "").Replace("\r", "").Trim();
                                string seoDescription = description.Replace("\n", "").Replace("\r", "").Trim();

                                SetProduct(xlWorkSheet, CurrentRow, name, slug, description, type, tags, price, imgUrl, seoTitle, seoDescription);

                                CurrentRow++;
                            }
                            catch (Exception e)
                            {
                                LogError("Error scraping product, skipping: " + e.Message);
                                continue;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        LogError("Error scraping page, skipping: " + ex.Message);
                        continue;
                    }
                }

                string excelPath = Path.Combine(MainDir, "products.csv");

                xlWorkBook.SaveAs(excelPath, Excel.XlFileFormat.xlCSV, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkBook.Close(true, misValue, misValue);
                excel.Quit();

                Marshal.ReleaseComObject(xlWorkSheet);
                Marshal.ReleaseComObject(xlWorkBook);
                Marshal.ReleaseComObject(excel);

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch (ThreadAbortException) { }
            catch (Exception ex)
            {
                LogError("Error scraping: " + ex.Message);
                CMBox.Show("Error", "Error scraping: " + ex.Message, Style.Error, Buttons.OK, ex.ToString());
            }
            finally
            {
                SetStatus("Idle");
                SetControlProperty(BtnStart, "Text", "Start");
                Log("Done" + newLine);
            }
        }
        #endregion

        #region "Handles"
        private void BtnSavePath_Click(object sender, EventArgs e)
        {
            VistaFolderBrowserDialog saveDialog = new VistaFolderBrowserDialog()
            {
                UseDescriptionForTitle = true,
                Description = "Select new save path",
                ShowNewFolderButton = true
            };

            if (saveDialog.ShowDialog() == DialogResult.OK)
            {
                TxtSavePath.Text = saveDialog.SelectedPath;
            }
        }

        private void BtnStart_Click(object sender, EventArgs e)
        {
            if (BtnStart.Text == "Start")
            {
                if (scraperThread != null)
                {
                    try
                    {
                        scraperThread.Abort();
                    }
                    catch { }

                    scraperThread = null;
                }

                scraperThread = new Thread(() => { Scrape(); }) { IsBackground = true };
                scraperThread.SetApartmentState(ApartmentState.STA);
                scraperThread.Start();

                SetStatus("Idle");
                BtnStart.Text = "Stop";
            }
            else
            {
                if (scraperThread != null)
                {
                    try
                    {
                        scraperThread.Abort();
                    }
                    catch { }

                    scraperThread = null;
                }

                BtnStart.Text = "Start";
            }
        }
        #endregion
    }
}