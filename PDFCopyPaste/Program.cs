using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;

namespace PDFCopyPaste
{
    class Form1 : Form
    {
        [System.Runtime.InteropServices.DllImport("user32")]
        private static extern IntPtr SetClipboardViewer(IntPtr hwnd);
        [System.Runtime.InteropServices.DllImport("user32")]
        private static extern IntPtr ChangeClipboardChain(IntPtr hwnd, IntPtr hWndNext);
        [System.Runtime.InteropServices.DllImport("user32")]
        private static extern int SendMessage(IntPtr hwnd, int wMsg, IntPtr wParam, IntPtr lParam);
        [System.Runtime.InteropServices.DllImport("user32")]
        private static extern IntPtr GetClipboardOwner();
        [System.Runtime.InteropServices.DllImport("user32")]
        private static extern int GetWindowThreadProcessId(IntPtr handle, out int processId);
        [System.Runtime.InteropServices.DllImport("user32")]
        private static extern bool CloseHandle(IntPtr handle);
        [System.Runtime.InteropServices.DllImport("kernel32")]
        public static extern int OpenProcess(int dwDesiredAccess, bool bInheritHandle, int dwProcessId);
        [DllImport("Psapi.dll", EntryPoint = "GetModuleFileNameEx")]
        public static extern uint GetModuleFileNameEx(int handle, int hModule, [Out] StringBuilder lpszFileName, uint nSize);

        const int WM_DRAWCLIPBOARD = 0x308;
        const int WM_CHANGECBCHAIN = 0x30D;

        public Form1()
        {
            Load += Form1_Load;
            Closed += Form1_Closed;
            Shown += Forms1_Shown;
            WindowState = FormWindowState.Minimized;
        }

        private void Forms1_Shown(object sender, EventArgs e)
        {
            Visible = false;
        }

        private string ProcessText(string str)
        {
            StringBuilder sb = new StringBuilder();
            int len = str.Length;
            int start = 0, end = str.IndexOf("\r\n");
            while (end != -1)
            {
                int tmp = end + 2;
                while (end > 0 && str[end - 1] == ' ') --end;
                if (str[end - 1] == '-') --end;
                sb.Append(str.Substring(start, end - start));
                if (str[end] != '-') sb.Append(' ');
                start = tmp;
                end = str.IndexOf("\r\n", start);
            }
            if (start != len)
                sb.Append(str.Substring(start, len - start));
            string res = Regex.Replace(sb.ToString(), @" *\([^\)]*20[0-9][0-9][^\)]*\)", "");
            return res;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //获得观察链中下一个窗口句柄
            NextClipHwnd = SetClipboardViewer(this.Handle);
        }

        private static object lck = new object();
        const int PROCESS_QUERY_INFORMATION = 0x0400;
        const int PROCESS_VM_READ = 0x0010;

        protected override void WndProc(ref System.Windows.Forms.Message m)
        {
            switch (m.Msg)
            {
                case WM_DRAWCLIPBOARD:
                    try
                    {
                        //检测文本
                        IDataObject obj = Clipboard.GetDataObject();
                        string[] formats = obj.GetFormats();
                        var owner = GetClipboardOwner();
                        int pid;
                        GetWindowThreadProcessId(owner, out pid);
                        int hProcess = OpenProcess(PROCESS_QUERY_INFORMATION | PROCESS_VM_READ, false, pid);
                        StringBuilder buffer = new StringBuilder();
                        buffer.EnsureCapacity(100);
                        GetModuleFileNameEx(hProcess, 0, buffer, 100);
                        if (buffer.ToString().IndexOf("PDF") != -1 && obj.GetDataPresent("System.String"))
                        {
                            string str = (string)obj.GetData("System.String");
                            string res = ProcessText(str);
                            if (res != str)
                            {
                                Thread.Sleep(500);
                                Clipboard.SetDataObject(res, true, 10000, 100);
                            }
                        }
                        else if (owner == IntPtr.Zero && obj.GetDataPresent("System.String"))
                        {
                            string str = (string)obj.GetData("System.String");
                            string res = Regex.Replace(str, "\r\n", "\n");
                            res = Regex.Replace(res, " +$", "", RegexOptions.Multiline);
                            if (res != str)
                            {
                                Thread.Sleep(500);
                                Clipboard.SetDataObject(res, true, 10000, 100);
                            }
                        }
                    } catch (Exception e)
                    {
                        MessageBox.Show($"{e.GetType()}: {e.Message}\nAborting...", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        Application.Exit();
                    }
                    //将WM_DRAWCLIPBOARD消息传递到下一个观察链中的窗口
                    SendMessage(NextClipHwnd, m.Msg, m.WParam, m.LParam);
                    break;
                default:
                    if (m.Msg == Program.WM_PROMPT_EXIT)
                    {
                        if (MessageBox.Show("Close PDFCopyPaste?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)
                            == DialogResult.Yes)
                        {
                            Application.Exit();
                        }
                    }
                    base.WndProc(ref m);
                    break;
            }
        }

        private void Form1_Closed(object sender, System.EventArgs e)
        {
            //从观察链中删除本观察窗口（第一个参数：将要删除的窗口的句柄；第二个参数：观察链中下一个窗口的句柄 ）
            ChangeClipboardChain(this.Handle, NextClipHwnd);
            //将变动消息WM_CHANGECBCHAIN消息传递到下一个观察链中的窗口
            SendMessage(NextClipHwnd, WM_CHANGECBCHAIN, this.Handle, NextClipHwnd);
        }

        IntPtr NextClipHwnd;
    }
    
    class Program
    {
        static Mutex m_mutex = new Mutex(false, "{BB20662C-D8CA-4E1B-A1B5-ABDC73D14926}");
        [DllImport("user32")]
        public static extern bool PostMessage(IntPtr hwnd, int msg, IntPtr wparam, IntPtr lparam);
        [DllImport("user32")]
        public static extern int RegisterWindowMessage(string message);
        public static int WM_PROMPT_EXIT = RegisterWindowMessage("WM_PROMPT_EXIT");
        [STAThread]
        static void Main(string[] args)
        {
            if (m_mutex.WaitOne(0, true))
            {
                Application.Run(new Form1());
                m_mutex.ReleaseMutex();
            } else
            {
                PostMessage(new IntPtr(0xffff), WM_PROMPT_EXIT, IntPtr.Zero, IntPtr.Zero);
            }
        }
    }
}
