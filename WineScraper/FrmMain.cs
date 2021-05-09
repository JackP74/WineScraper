using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WineScraper
{
    public partial class FrmMain : Form
    {
        #region "Variables"
        private readonly string AppPath = Application.StartupPath;
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

        #region "Functions"
        public FrmMain()
        {
            InitializeComponent();

            CreateConsole();
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

        private void Scrape()
        {
            using WebClient client = new WebClient();
            client.Headers.Add("user-agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:88.0) Gecko/20100101 Firefox/88.0");


        }
        #endregion
    }
}
