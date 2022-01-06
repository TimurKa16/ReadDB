// Программа для считывания таблицы, которую сложно самому копировать
// Задаются координаты ячеек одной строки + в конце точка для прокрутки
// Таким образом, считывается вся таблица

// F10 - задание координат ячеек
// F11 - начать считывание
// F12 - остановить считывание, запись в файл







using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Collections.Generic;
using System.IO;


namespace ReadDB
{
    public partial class Form1 : Form
    {
        List<string> transactions = new List<string>();

        string fileToWrite = @"File.txt";
        List<POINT> points = new List<POINT>();


        public Form1()
        {
            InitializeComponent();
            SetHook();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        POINT p1 = new POINT();
        void AddPoint()
        {
            POINT p = new POINT();
            GetCursorPos(ref p);
            points.Add(p);
        }


        public void StartReading()
        {
            ReadTable(ref transactions);
            WriteToFile(fileToWrite, transactions);
        }

        void StopReading()
        {
            //p1 = new POINT();
            points = new List<POINT>();
        }


        

        private void ReadTable(ref List<string> transactions)
        {
            for (int i = 0; i < 1; i++)
            {

                transactions.Add(ReadRaw(points));
            }
        }

        private string ReadRaw(List<POINT> points)
        {
            string raw = null;

            for (int i = 0; i < points.Count - 1; i++)
                raw += ReadCell(points[i].x, points[i].y) + "\\t";

            ClickMouse(points[points.Count - 1].x, points[points.Count - 1].y);

            return raw;
        }

        private string ReadCell(int x, int y)
        {
            ClickMouse(x, y);
            SendKeys.Send("^{c}");
            return Clipboard.GetText();
        }

        public static void WriteToFile(string fileToWrite, List<string> transaction)
        {
            using (StreamWriter streamWriter = new StreamWriter(fileToWrite))
            {
                for (int i = 0; i < transaction.Count; i++)
                    streamWriter.WriteLine(transaction[i]);
            }
        }


        //__________________________________________________________________________________________________________________________________//
        //__________________________________________________________________________________________________________________________________//












        //__________________________________________________________________________________________________________________________________//
        //__________________________________________________________________________________________________________________________________//

        private void ClickMouse(int x, int y)
        {
            SetCursorPos(x, y);
            DoMouseLeftClick(x, y);
        }


        [DllImport("user32.dll")]
        public static extern bool ClientToScreen(IntPtr hWind, ref POINT point);

        [StructLayout(LayoutKind.Sequential)]
        public struct POINT
        {
            public int x;
            public int y;
        }


        [DllImport("user32.dll")]
        public static extern void mouse_event(int dsFlags, int dx, int dy, int cButtons, int dsExtraInfo);

        public const int MOUSEEVENTF_LEFTDOWN = 0x02;
        public const int MOUSEEVENTF_LEFTUP = 0x04;

        public const int MOUSEEVENTF_RIGHTDOWN = 0x08;
        public const int MOUSEEVENTF_RIGHTUP = 0x10;

        private void DoMouseLeftClick(int x, int y)
        {
            mouse_event(MOUSEEVENTF_LEFTDOWN, x, y, 0, 0);
            mouse_event(MOUSEEVENTF_LEFTUP, x, y, 0, 0);
        }

        [DllImport("user32.dll")]
        public static extern bool GetCursorPos(ref POINT lpPoint);

        [DllImport("user32.dll")]
        public static extern bool SetCursorPos(int x, int y);


        //__________________________________________________________________________________________________________________________________//
        //__________________________________________________________________________________________________________________________________//

        private const int WH_KEYBOARD_LL = 13;


        private static LowLevelKeyboardProcDelegate m_callback;
        private static IntPtr m_hHook;

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(
            int idHook,
            LowLevelKeyboardProcDelegate lpfn,
            IntPtr hMod, int dwThreadId);


        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);


        [DllImport("Kernel32.dll", SetLastError = true)]
        private static extern IntPtr GetModuleHandle(IntPtr lpModuleName);


        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr CallNextHookEx(
            IntPtr hhk,
            int nCode, IntPtr wParam, IntPtr lParam);

        private IntPtr LowLevelKeyboardHookProc(
            int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0)
            {
                var khs = (KeyboardHookStruct)
                          Marshal.PtrToStructure(lParam,
                          typeof(KeyboardHookStruct));
                if (Convert.ToInt32("" + wParam) == 256)
                {
                    //MessageBox.Show(khs.VirtualKeyCode+""); //Показать Ключ нажатой клавиши
                    if ((int)khs.VirtualKeyCode == 121)//121 - F10
                    {
                        AddPoint();

                    }
                    else if ((int)khs.VirtualKeyCode == 122)//122 - F11
                    {
                        StartReading();

                    }
                    else if ((int)khs.VirtualKeyCode == 123)//123 - F12
                    {
                        StopReading();

                    }
                }
            }
            return CallNextHookEx(m_hHook, nCode, wParam, lParam);
        }


        [StructLayout(LayoutKind.Sequential)]
        private struct KeyboardHookStruct
        {
            public readonly int VirtualKeyCode;
            public readonly int ScanCode;
            public readonly int Flags;
            public readonly int Time;
            public readonly IntPtr ExtraInfo;
        }


        private delegate IntPtr LowLevelKeyboardProcDelegate(
            int nCode, IntPtr wParam, IntPtr lParam);


        public void SetHook()
        {
            m_callback = LowLevelKeyboardHookProc;
            m_hHook = SetWindowsHookEx(WH_KEYBOARD_LL,
                m_callback,
                GetModuleHandle(IntPtr.Zero), 0);
        }

        public static void Unhook()
        {
            UnhookWindowsHookEx(m_hHook);
        }
    }
}

