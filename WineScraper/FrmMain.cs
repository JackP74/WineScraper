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

namespace WineScraper
{
    public partial class FrmMain : Form
    {
        #region "Variables"
        private readonly string AppPath = Application.StartupPath;
        private readonly string WineUrl = @"https://www.maccaninodrink.com/en-gb/cameras/?limit=100";
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
                client.Headers.Add("user-agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:88.0) Gecko/20100101 Firefox/88.0");

                string rawHtml = client.DownloadString(WineUrl);

                if (string.IsNullOrWhiteSpace(rawHtml))
                {
                    LogError("No response from server");
                    return;
                }

                var doc = new HtmlDocument();
                doc.LoadHtml(rawHtml);

                Log("Getting pages...");

                List<string> productsURLs = new List<string>();

                var parentElement = doc.GetElementbyId("content");
                var productsDivs = parentElement.SelectNodes(".//div[contains(@class, 'product-thumb')]");

                foreach(var product in productsDivs)
                {
                    string url = product.SelectSingleNode(".//div[@class='name']").SelectSingleNode(".//a[@href]").GetAttributeValue("href", string.Empty);

                    if (url.EndsWith("?limit=100"))
                        url = url.Substring(0, url.Length - 10);

                    bool result = Uri.TryCreate(url, UriKind.Absolute, out Uri uriResult) && (uriResult.Scheme == Uri.UriSchemeHttp || uriResult.Scheme == Uri.UriSchemeHttps);

                    if (result)
                        productsURLs.Add(url);
                }

                Log("Getting products...");

                foreach(var url in productsURLs)
                {
                    rawHtml = client.DownloadString(url);

                    doc = new HtmlDocument();
                    doc.LoadHtml(rawHtml);

                    parentElement = doc.GetElementbyId("content");
                    var productDiv = parentElement.SelectSingleNode(".//div[@class='row']").SelectSingleNode(".//div[@class='product-buy-wrapper']");

                    string name = productDiv.SelectSingleNode(".//h1").InnerText;

                    string validFileName = MakeValidFileName(name);
                    string productPath = Path.Combine(MainDir, validFileName);

                    if (!Directory.Exists(productPath))
                        Directory.CreateDirectory(productPath);

                    var imgDiv = doc.GetElementbyId("zoom1");
                    string imgUrl = imgDiv.GetAttributeValue("href", string.Empty);
                    imgUrl = imgUrl.Replace("image/cache/catalog", "image/catalog");
                    imgUrl = imgUrl.Substring(0, imgUrl.Length - 12) + ".jpg";

                    string imgFileName = Path.Combine(productPath, MakeValidFileName(imgUrl.Split('/').Last()));

                    client.DownloadFile(imgUrl, imgFileName);


                }

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