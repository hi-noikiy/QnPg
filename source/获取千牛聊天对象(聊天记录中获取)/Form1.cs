using MasterDevs.ChromeDevTools;
using MasterDevs.ChromeDevTools.Protocol.Chrome.DOM;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace 获取千牛聊天对象
{
    public partial class Form1 : Form
    {
        private string chatWindowTitlePattern = ".*(?= - 接待中心)";

        public Dictionary<string,int> QnAccounts { get; set; }

        public HashSet<string> ChatlogChromeTitles { get; set; }

        Dictionary<string, ChromeSession> ChromeSessions;

        public ChromeSession CurrentChromeSession { get; set; }

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.ChatlogChromeTitles = new HashSet<string>
			{
				"当前聊天窗口",
				"IMKIT.CLIENT.QIANNIU",
				"聊天窗口",
				"imkit.qianniu"
			};

            ChromeSessions = new Dictionary<string, ChromeSession>();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            QnAccounts = GetAllChatDeskSellerNameAndHwndInner();
            cmbQnAccount.Items.AddRange(QnAccounts.Keys.ToArray());
        }

        private void cmbQnAccount_SelectedIndexChanged(object sender, EventArgs e)
        {
            var selectedItem = cmbQnAccount.Items[cmbQnAccount.SelectedIndex].ToString();
            var qnHwnd = QnAccounts.First(k=>k.Key==selectedItem).Value;
            var sessionInfoUrl = GetSessionInfoUrl(qnHwnd);
            if (!ChromeSessions.ContainsKey(selectedItem))
            {
                ChromeSessions[selectedItem] = CreateChromeSession(sessionInfoUrl, ChatlogChromeTitles);
            }
            CurrentChromeSession = ChromeSessions[selectedItem];
        }

        /// <summary>
        /// 获取当前聊天的对象
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void timer1_Tick(object sender, EventArgs e)
        {
            if (cmbQnAccount.SelectedIndex < 0) return;
            txtGuestName.Text = GetGuestNameFromChatRecord();
        }

        #region 解析聊天记录的html

        public string GetGuestNameFromChatRecord() 
        {
            string guestName = string.Empty;

            if (cmbQnAccount.SelectedIndex > -1)
            {
                var selectedItem = cmbQnAccount.Items[cmbQnAccount.SelectedIndex].ToString();
                var html = string.Empty;
                if (GetHtml(out html))
                {
                    var parser = new ChatRecordParser(selectedItem);
                    var chatRecord = parser.Parse(html);
                    guestName = chatRecord==null ? string.Empty : chatRecord.GuestName;
                    //chatRecord.Statements.ForEach(k =>
                    //{
                    //    Console.WriteLine("speaker={0},time={1},content={2}", k.Speaker, k.Time, k.OneStringWithTrim);
                    //});
                }
            }

            return guestName;
        }

        #endregion


        public bool GetHtml(out string html)
        {
            bool hasGetHtml = false;
            html = "";
            ICommandResponse commandResponse;
            if (this.SendCommandSafe<GetDocumentCommand>(out commandResponse, null))
            {
                CommandResponse<GetDocumentCommandResponse> commandResponse2 = commandResponse as CommandResponse<GetDocumentCommandResponse>;
                long nodeId = commandResponse2.Result.Root.NodeId;
                GetOuterHTMLCommand parameter = new GetOuterHTMLCommand
                {
                    NodeId = nodeId
                };
                if (this.SendCommandSafe<GetOuterHTMLCommand>(out commandResponse, parameter))
                {
                    CommandResponse<GetOuterHTMLCommandResponse> commandResponse3 = commandResponse as CommandResponse<GetOuterHTMLCommandResponse>;
                    html = (commandResponse3.Result.OuterHTML ?? "");
                    hasGetHtml = true;
                }
            }
            return hasGetHtml;
        }

        public bool SendCommandSafe<T>(out ICommandResponse rsp, T parameter = default(T)) where T : class
        {
            rsp = null;
            bool result = true;
            try
            {
                if (this.CurrentChromeSession == null)
                {
                    throw new Exception("Session=Null,parameter=" + parameter.ToString());
                }
                rsp = this.CurrentChromeSession.SendCommand<T>(parameter);
                if (rsp == null)
                {
                    throw new Exception("SendCommand返回null");
                }
                if (rsp is IErrorResponse)
                {
                    IErrorResponse errorResponse = rsp as IErrorResponse;
                    throw new Exception(string.Concat(new object[]
					{
						"SendCommand error,msg=",
						errorResponse.Error.Message,
						",code=",
						errorResponse.Error.Code
					}));
                }
            }
            catch (Exception ex)
            {
                result = false;
            }
            return result;
        }


        private ChromeSession CreateChromeSession(string sessionInfoUrl, HashSet<string> titles)
        {
            var sessionInfos = GetWebSocketSessionInfos(sessionInfoUrl);
            sessionInfos = sessionInfos.Where(k => titles.Contains(k.Title)).ToList();
            if (sessionInfos == null || sessionInfos.Count < 1) throw new Exception("");

            if (sessionInfos.Count(k => string.IsNullOrEmpty(k.WebSocketDebuggerUrl)) > 0)
            {
                throw new Exception("");
            }
            var webSocketSessionInfo = sessionInfos[0];
            if (sessionInfos.Count > 1)
            {
                webSocketSessionInfo = sessionInfos.First(k => !k.Url.ToLower().Contains("type=1&"));
            }
            return new ChromeSessionFactory().Create(webSocketSessionInfo.WebSocketDebuggerUrl, webSocketSessionInfo.Title);
        }

        private static List<WebSocketSessionInfo> GetWebSocketSessionInfos(string endpointUrl)
        {
            new List<WebSocketSessionInfo>();
            List<WebSocketSessionInfo> sessionInfos = null;
            using (WebClient webClient = new WebClient())
            {
                UriBuilder uriBuilder = new UriBuilder(endpointUrl);
                uriBuilder.Path = "/json";
                webClient.Encoding = Encoding.UTF8;
                string value = webClient.DownloadString(uriBuilder.Uri);
                sessionInfos = JsonConvert.DeserializeObject<List<WebSocketSessionInfo>>(value);
            }
            return sessionInfos;
        }

        public static string GetSessionInfoUrl(int hwnd)
        {
            if (hwnd == 0)
            {
                throw new Exception("GetSession,hwnd=0");
            }
            int pid = GetWindowThreadProcessId(hwnd);
            if (pid == 0)
            {
                throw new Exception("GetSession,pid=0");
            }
            int port = GetChromeListeningPortWithInterOp(pid);
            if (port == 0)
            {
                throw new Exception("GetSession,port=0");
            }
            return string.Format("http://localhost:{0}", port);
        }

        private static int GetChromeListeningPortWithInterOp(int pid)
        {
            int port = 0;
            if (pid > 0)
            {
                port = TcpConnectionInfo.GetChromeListeningPortWithInterOp(pid);
            }
            return port;
        }

        [DllImport("user32.dll")]
        private static extern int GetWindowThreadProcessId(int hWnd, ref int lpdwProcessId);

        private static int GetWindowThreadProcessId(int hwnd)
        {
            int lpdwProcessId = 0;
            GetWindowThreadProcessId(hwnd, ref lpdwProcessId);
            return lpdwProcessId;
        }

        #region [WIMAPI部分]

        public virtual Dictionary<string, int> GetAllChatDeskSellerNameAndHwndInner()
        {
            Dictionary<string, int> rtdict = new Dictionary<string, int>();
            FindAllDesktopWindowByClassNameAndTitlePattern("StandardFrame", chatWindowTitlePattern, (int hwnd, string title) =>
            {
                if (IsWindowVisible(hwnd))
                {
                    string winTitle = Regex.Match(title, chatWindowTitlePattern).ToString();
                    if (!string.IsNullOrEmpty(winTitle))
                    {
                        rtdict[winTitle] = hwnd;
                    }
                }
            });
            return rtdict;
        }

        public List<int> FindAllDesktopWindowByClassNameAndTitlePattern(string cname, string pattern = null, Action<int, string> action = null)
        {
            List<int> rtlist = new List<int>();
            TraverAllDesktopHwnd((int hwnd) =>
            {
                bool isMatch = true;
                string title = null;
                if (!string.IsNullOrEmpty(pattern))
                {
                    isMatch = (GetText(hwnd, out title) && Regex.IsMatch(title, pattern));
                }
                if (isMatch)
                {
                    rtlist.Add(hwnd);
                    if (action != null)
                    {
                        action(hwnd, title);
                    }
                }
            }, cname, null);
            return rtlist;
        }

        public static void TraverAllDesktopHwnd(Action<int> action, string className = null, string title = null)
        {
            TraverseChildHwnd(0, action, false, className, title);
        }

        public static void TraverseChildHwnd(int parent, Action<int> act, bool traverseDescandent, string className = null, string title = null)
        {
            for (int i = FindWindowEx(parent, 0, className, title); i > 0; i = FindWindowEx(parent, i, className, title))
            {
                act(i);
                if (traverseDescandent)
                {
                    TraverseChildHwnd(i, act, traverseDescandent, className, title);
                }
            }
        }

        public static bool GetText(int hWnd, out string txt)
        {
            txt = "";
            StringBuilder stringBuilder = new StringBuilder(8192);
            bool result;
            if (result = SendForGetText(hWnd, stringBuilder, 2000))
            {
                txt = stringBuilder.ToString();
            }
            return result;
        }

        public static bool SendForGetText(int hWnd, StringBuilder sbText, int uTimeout = 2000)
        {
            bool result = false;
            int lpdwResult;
            result = SendMessageTimeout(hWnd, 13, sbText.Capacity, sbText, SendMessageTimeoutFlags.SMTO_ABORTIFHUNG | SendMessageTimeoutFlags.SMTO_ERRORONEXIT, uTimeout, out lpdwResult);
            return result;
        }


        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern bool SendMessageTimeout(int hWnd, int Msg, int wParam, int lParam, SendMessageTimeoutFlags fuFlags, int uTimeout, out int lpdwResult);
        [DllImport("user32.dll", CharSet = CharSet.Auto, EntryPoint = "SendMessageTimeout", SetLastError = true)]
        public static extern bool SendMessageTimeout(int hWnd, int Msg, int wParam, StringBuilder lParam, SendMessageTimeoutFlags fuFlags, int uTimeout, out int lpdwResult);
        [DllImport("user32.dll")]
        public static extern int FindWindowEx(int hwndParent, int hwndChildAfter, string lpszClass, string lpszWindow);
        [DllImport("user32.dll")]
        public static extern bool IsWindowVisible(int hWnd);

        [Flags]
        public enum SendMessageTimeoutFlags : uint
        {
            SMTO_NORMAL = 0u,
            SMTO_BLOCK = 1u,
            SMTO_ABORTIFHUNG = 2u,
            SMTO_NOTIMEOUTIFNOTHUNG = 8u,
            SMTO_ERRORONEXIT = 32u
        }
        #endregion


    }

    public class WebSocketSessionInfo
    {
        public string Description;
        public string DevtoolsFrontendUrl;
        public string Id;
        public string Title;
        public string Type;
        public string Url;
        public string WebSocketDebuggerUrl;
        public WebSocketSessionInfo()
        {
        }
    }

    public class TcpConnectionInfo
    {
        public enum TCP_TABLE_CLASS : int
        {
            TCP_TABLE_BASIC_LISTENER,
            TCP_TABLE_BASIC_CONNECTIONS,
            TCP_TABLE_BASIC_ALL,
            TCP_TABLE_OWNER_PID_LISTENER,
            TCP_TABLE_OWNER_PID_CONNECTIONS,
            TCP_TABLE_OWNER_PID_ALL,
            TCP_TABLE_OWNER_MODULE_LISTENER,
            TCP_TABLE_OWNER_MODULE_CONNECTIONS,
            TCP_TABLE_OWNER_MODULE_ALL
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct MIB_TCPROW_OWNER_PID
        {
            public uint state;
            public uint localAddr;
            public byte localPort1;
            public byte localPort2;
            public byte localPort3;
            public byte localPort4;
            public uint remoteAddr;
            public byte remotePort1;
            public byte remotePort2;
            public byte remotePort3;
            public byte remotePort4;
            public int owningPid;

            public ushort LocalPort
            {
                get
                {
                    return BitConverter.ToUInt16(
                        new byte[2] { localPort2, localPort1 }, 0);
                }
            }

            public ushort RemotePort
            {
                get
                {
                    return BitConverter.ToUInt16(
                        new byte[2] { remotePort2, remotePort1 }, 0);
                }
            }
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct MIB_TCPTABLE_OWNER_PID
        {
            public uint dwNumEntries;
            MIB_TCPROW_OWNER_PID table;
        }

        [DllImport("iphlpapi.dll", SetLastError = true)]
        static extern uint GetExtendedTcpTable(IntPtr pTcpTable,
            ref int dwOutBufLen,
            bool sort,
            int ipVersion,
            TCP_TABLE_CLASS tblClass,
            int reserved);

        public static MIB_TCPROW_OWNER_PID[] GetAllTcpConnections()
        {
            MIB_TCPROW_OWNER_PID[] tTable;
            int AF_INET = 2;    // IP_v4
            int buffSize = 0;

            // how much memory do we need?
            uint ret = GetExtendedTcpTable(IntPtr.Zero,
                ref buffSize,
                true,
                AF_INET,
                TCP_TABLE_CLASS.TCP_TABLE_OWNER_PID_ALL,
                0);
            if (ret != 0 && ret != 122) // 122 insufficient buffer size
                throw new Exception("bad ret on check " + ret);
            IntPtr buffTable = Marshal.AllocHGlobal(buffSize);

            try
            {
                ret = GetExtendedTcpTable(buffTable,
                    ref buffSize,
                    true,
                    AF_INET,
                    TCP_TABLE_CLASS.TCP_TABLE_OWNER_PID_ALL,
                    0);
                if (ret != 0)
                    throw new Exception("bad ret " + ret);

                // get the number of entries in the table
                MIB_TCPTABLE_OWNER_PID tab =
                    (MIB_TCPTABLE_OWNER_PID)Marshal.PtrToStructure(
                        buffTable,
                        typeof(MIB_TCPTABLE_OWNER_PID));
                IntPtr rowPtr = (IntPtr)((long)buffTable +
                    Marshal.SizeOf(tab.dwNumEntries));
                tTable = new MIB_TCPROW_OWNER_PID[tab.dwNumEntries];

                for (int i = 0; i < tab.dwNumEntries; i++)
                {
                    MIB_TCPROW_OWNER_PID tcpRow = (MIB_TCPROW_OWNER_PID)Marshal
                        .PtrToStructure(rowPtr, typeof(MIB_TCPROW_OWNER_PID));
                    tTable[i] = tcpRow;
                    // next entry
                    rowPtr = (IntPtr)((long)rowPtr + Marshal.SizeOf(tcpRow));
                }
            }
            finally
            {
                // Free the Memory
                Marshal.FreeHGlobal(buffTable);
            }
            return tTable;
        }

        public static int GetChromeListeningPortWithInterOp(int pid)
        {
            int port = 0;
            try
            {
                TcpConnectionInfo.MIB_TCPROW_OWNER_PID[] tcpConnections = TcpConnectionInfo.GetAllTcpConnections();
                for (int i = 0; i < tcpConnections.Length; i++)
                {
                    TcpConnectionInfo.MIB_TCPROW_OWNER_PID mIB_TCPROW_OWNER_PID = tcpConnections[i];
                    if (mIB_TCPROW_OWNER_PID.localAddr == 16777343u && mIB_TCPROW_OWNER_PID.owningPid == pid)
                    {
                        port = (int)mIB_TCPROW_OWNER_PID.LocalPort;
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                Trace.Assert(false, ex.Message);
            }
            return port;
        }
    }
}
