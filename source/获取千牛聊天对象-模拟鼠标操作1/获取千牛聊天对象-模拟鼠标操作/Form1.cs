
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace 获取千牛聊天对象_模拟鼠标操作
{
    public partial class Form1 : Form
    {
        private string chatWindowTitlePattern = ".*(?= - 接待中心)";

        public Dictionary<string, int> QnAccounts { get; set; }

        public int CurrentQnHwnd { get; set; }

        private string _taskBuyerPt;

        private Dictionary<string, string> _bmpB64_buyerNameDict;

        public DateTime _getGuestNameImageB64CacheTime { get; set; }

        public string _getImageB64StrCache { get; set; }

        public string _preB64 { get; set; }

        public Form1()
        {
            _bmpB64_buyerNameDict = new Dictionary<string, string>();
            this._taskBuyerPt = "(?<=^benchwebctrl:.*_bench_CQNLightTaskBiz_qtask_create_).*";
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        private void button1_Click(object sender, EventArgs e)
        {
            QnAccounts = GetAllChatDeskSellerNameAndHwndInner();
            cmbQnAccount.Items.AddRange(QnAccounts.Keys.ToArray());
        }

        private void cmbQnAccount_SelectedIndexChanged(object sender, EventArgs e)
        {
            var selectedItem = cmbQnAccount.Items[cmbQnAccount.SelectedIndex].ToString();
            CurrentQnHwnd = QnAccounts.First(k => k.Key == selectedItem).Value;
        }


        private void timer1_Tick(object sender, EventArgs e)
        {
            MainLoop();
        }

        private void MainLoop()
        {
            if (this.CurrentQnHwnd > 0)
            {
                txtGuestName.Text = GetBuyerNameUseCache();
            }
        }

        private string GetBuyerNameUseCache()
        {
            string buyerName = "";
            string buyerImageB64Str = this.GetImageB64StrUseCache(true);
            if (this._bmpB64_buyerNameDict.ContainsKey(buyerImageB64Str))
            {
                buyerName = this._bmpB64_buyerNameDict[buyerImageB64Str];
            }
            else
            {
                buyerName = this.GetBuyerName();
                if (!string.IsNullOrEmpty(buyerName))
                {
                    string buyerImageB64StrNoCache = this.GetImageB64StrUseCache(false);
                    if (buyerImageB64StrNoCache == buyerImageB64Str)
                    {
                        this._bmpB64_buyerNameDict[buyerImageB64Str] = buyerName;
                    }
                }
            }
            return buyerName;
        }

        private string GetImageB64StrUseCache(bool noCache)
        {
            if (!noCache || (DateTime.Now - this._getGuestNameImageB64CacheTime).TotalMilliseconds > 10.0)
            {
                Rectangle buyerNameRegion = this.GetBuyerNameRegion();
                this._getImageB64StrCache = this.GetImageB64Str(buyerNameRegion);
                this._getGuestNameImageB64CacheTime = DateTime.Now;
            }
            return this._getImageB64StrCache;
        }

        private string GetImageB64Str(Rectangle buyerNameRegion)
        {
            string newB64 = "";
            try
            {
                Bitmap buyerNameRegionImage = this.DrawToBitmap(buyerNameRegion);
                newB64 = this.BitmapToBase64String(buyerNameRegionImage);
                if (this._preB64 != newB64)
                {
                    this._preB64 = newB64;
                }
            }
            catch (Exception e)
            {
            }
            return newB64;
        }

        private string BitmapToBase64String(Bitmap bitmap)
        {
            string b64Str = "";
            using (MemoryStream memoryStream = new MemoryStream())
            {
                bitmap.Save(memoryStream, ImageFormat.Png);
                b64Str = Convert.ToBase64String(memoryStream.ToArray());
            }
            return b64Str;
        }
        private Bitmap DrawToBitmap(Rectangle rect)
        {
            Bitmap bitmap = new Bitmap(rect.Width, rect.Height);
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.CopyFromScreen(rect.Location, new Point(0, 0), rect.Size);
            }
            return bitmap;
        }

        public string GetBuyerName()
        {
            string buyerName = "";
            List<int> taskWinHwnds;
            int taskHwnd = this.TryGetTaskWindow(out taskWinHwnds);
            if (taskHwnd > 0)
            {
                buyerName = this.GetBuyerNameFromTaskWindow(taskHwnd);
                CloseWindow(taskHwnd, 2000);
            }
            if (buyerName != null)
            {
                buyerName = buyerName.Trim();
                string endStr = "(阿里巴巴中国站)";
                if (buyerName.EndsWith(endStr))
                {
                    buyerName = buyerName.Substring(0, buyerName.Length - endStr.Length);
                }
            }
            return buyerName;
        }

        private string GetBuyerNameFromTaskWindow(int pHwnd)
        {
            string buyerName = null;
            for (int i = 0; i < 4; i++)
            {
                int hwnd = FindDescendantHwnd(pHwnd, this.GetTaskWindowClueDontUse(), "GetBuyerNameFromTaskWindow");
                if (hwnd > 0)
                {
                    string text = GetText(hwnd);
                    if (!string.IsNullOrEmpty(text))
                    {
                        Match match = Regex.Match(text, this._taskBuyerPt);
                        if (match != null && match.Success)
                        {
                            buyerName = match.ToString().Trim();
                        }
                    }
                    else
                    {
                    }
                    return buyerName;
                }
                Thread.Sleep(200);
            }
            return buyerName;
        }


        private int TryGetTaskWindow(out List<int> taskWinHwnds)
        {
            taskWinHwnds = this.GetAllTaskWindow();
            this.ClickTask();
            int loopNum = 0;
            int hWnd;
            do
            {
                hWnd = this.GetTopTaskWindow(taskWinHwnds);
                if (hWnd == 0)
                {
                    Thread.Sleep(20);
                    loopNum++;
                }
            }
            while (hWnd == 0 && loopNum < 5);

            return hWnd;
        }

        private int GetTopTaskWindow(List<int> taskWinHwnds)
        {
            int hWnd = 0;
            List<int> newTaskWinHwnds = this.GetAllTaskWindow();
            if (newTaskWinHwnds.Count > 0)
            {
                foreach (int newHwnd in newTaskWinHwnds)
                {
                    if (!taskWinHwnds.Contains(newHwnd))
                    {
                        hWnd = newHwnd;
                        break;
                    }
                }
            }
            return hWnd;
        }

        private void ClickTask()
        {
            try
            {
                int toolbarPlusHwnd = this.ToolbarPlusHwnd;
                Rectangle windowRectangle = GetWindowRectangle(toolbarPlusHwnd);
                //添加任务按钮坐标
                int x = 125; 
                int y = 18;
                ClickPointBySendMessage(toolbarPlusHwnd, x, y);
            }
            catch (Exception e)
            {
            }
        }
        private List<int> GetAllTaskWindow()
        {
            List<int> hwnds = new List<int>();
            TraverAllDesktopHwnd(hWnd =>
            {
                if (this.IsTaskWindow(hWnd))
                {
                    hwnds.Add(hWnd);
                }
            }, "StandardFrame", "添加任务");
            return hwnds;
        }
        private bool IsTaskWindow(int hwnd)
        {
            return FindDescendantHwnd(hwnd, this.GetTaskWindowClueDontUse(), "IsTaskWindow") > 0;
        }

        private List<WindowClue> _toolbarPlusClueDontUse;
        private List<WindowClue> GetToolbarPlusClueDontUse()
        {
            if (this._toolbarPlusClueDontUse == null)
            {
                this._toolbarPlusClueDontUse = this.GetToolbarPlusClueUnCache();
            }
            return this._toolbarPlusClueDontUse;
        }

        protected virtual List<WindowClue> GetToolbarPlusClueUnCache()
        {
            return new List<WindowClue>
			{
				new WindowClue(ClsNameEnum.StandardWindow, null, -1),
				new WindowClue(ClsNameEnum.StandardWindow, null, 1),
				new WindowClue(ClsNameEnum.ToolBarPlus, null, -1)
			};
        }
        private int _toolbarPlusHwnd;
        public int ToolbarPlusHwnd
        {
            get
            {
                if (this._toolbarPlusHwnd == 0)
                {
                    this._toolbarPlusHwnd = FindDescendantHwnd(this.CurrentQnHwnd, this.GetToolbarPlusClueDontUse(), "ToolbarPlusHwnd");
                }
                return this._toolbarPlusHwnd;
            }
        }

        private List<WindowClue> _taskWindowClueDontUse;
        private List<WindowClue> GetTaskWindowClueDontUse()
        {
            if (this._taskWindowClueDontUse == null)
            {
                this._taskWindowClueDontUse = new List<WindowClue>();
                this._taskWindowClueDontUse.Add(new WindowClue(ClsNameEnum.PrivateWebCtrl, null, -1));
                this._taskWindowClueDontUse.Add(new WindowClue(ClsNameEnum.Aef_WidgetWin_0, null, -1));
            }
            return this._taskWindowClueDontUse;
        }

        private int _buyerPicHwnd;
        public int BuyerPicHwnd
        {
            get
            {
                if (this._buyerPicHwnd == 0)
                {
                    this._buyerPicHwnd = FindDescendantHwnd(CurrentQnHwnd, this.GetBuyerPicClue(), "BuyerPicHwnd");
                }
                return this._buyerPicHwnd;
            }
        }

        private List<WindowClue> _buyerPicClue;
        private List<WindowClue> GetBuyerPicClue()
        {
            if (this._buyerPicClue == null)
            {
                this._buyerPicClue = this.GetBuyerPicClueUnCache();
            }
            return this._buyerPicClue;
        }
        protected virtual List<WindowClue> GetBuyerPicClueUnCache()
        {
            return new List<WindowClue>
			{
				new WindowClue(ClsNameEnum.StandardWindow, null, -1),
				new WindowClue(ClsNameEnum.StandardWindow, null, 1),
				new WindowClue(ClsNameEnum.StandardWindow, null, 1),
				new WindowClue(ClsNameEnum.StandardButton, null, -1)
			};
        }

        public Rectangle GetBuyerNameRegion()
        {
            int buyerPicHwnd = this.BuyerPicHwnd;
            Rectangle windowRectangle = GetWindowRectangle(buyerPicHwnd);
            return new Rectangle(windowRectangle.Right + 8, windowRectangle.Top + 10, 200, 3);
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

        public static string GetText(int hwnd)
        {
            string text;
            if (!GetText(hwnd, out text))
            {
                text = null;
            }
            return text;
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


        public static int FindDescendantHwnd(int parentHwnd, List<WindowClue> clueList, [System.Runtime.CompilerServices.CallerMemberName] string callerName = "")
        {
            int hWnd = FindDescendantHwnd(parentHwnd, clueList, 0);
            if (hWnd == 0)
            {
            }
            return hWnd;
        }

        private static int FindDescendantHwnd(int parentHwnd, List<WindowClue> clueList, int startIdx)
        {
            if (clueList == null || clueList.Count == 0)
            {
                throw new Exception("没有WindowClue");
            }
            int hwndChild = 0;
            WindowClue windowClue = clueList[startIdx];
            int skipCnt = (windowClue.SkipCount > 0) ? windowClue.SkipCount : 0;
            for (int i = 0; i <= skipCnt; i++)
            {
                hwndChild = FindWindowEx(parentHwnd, hwndChild, windowClue.ClsName, windowClue.Text);
                if (hwndChild == 0)
                {
                    return hwndChild;
                }
            }
            if (startIdx < clueList.Count - 1)
            {
                int dehWnd = FindDescendantHwnd(hwndChild, clueList, startIdx + 1);
                if (dehWnd == 0)
                {
                    while (dehWnd == 0)
                    {
                        hwndChild = FindWindowEx(parentHwnd, hwndChild, windowClue.ClsName, windowClue.Text);
                        if (hwndChild == 0)
                        {
                            return hwndChild;
                        }
                        dehWnd = FindDescendantHwnd(hwndChild, clueList, startIdx + 1);
                    }
                    hwndChild = dehWnd;
                }
                else
                {
                    hwndChild = dehWnd;
                }
            }
            return hwndChild;
        }

        public static void ClickPointBySendMessage(int hwnd, int x, int y)
        {
            if (x >= 0 && y >= 0)
            {
                y <<= 16;
                y = (x | y);
                SendMessage(hwnd, 513, 1, y, 2000, "ClickPointBySendMessage");
                SendMessage(hwnd, 514, 0, y, 2000, "ClickPointBySendMessage");
            }
        }

        public static bool CloseWindow(int hwnd, int timeoutMs = 2000)
        {
            return SendMessage(hwnd, 16, 0, 0, timeoutMs, "CloseWindow");
        }

        public static bool SendMessage(int hWnd, int msg, int wParam, int lParam, int uTimeout = 2000, [System.Runtime.CompilerServices.CallerMemberName] string caller = "")
        {
            int lpdwResult;
            return SendMessage(hWnd, caller, msg, wParam, lParam, out lpdwResult, uTimeout);
        }

        public static bool SendMessage(int hWnd, string caller, int msg, int wParam, int lParam, out int lpdwResult, int uTimeout = 2000)
        {
            lpdwResult = 0;
            return SendMessageTimeout(hWnd, msg, wParam, lParam, SendMessageTimeoutFlags.SMTO_ABORTIFHUNG | SendMessageTimeoutFlags.SMTO_ERRORONEXIT, uTimeout, out lpdwResult);
        }

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern bool SendMessageTimeout(int hWnd, int Msg, int wParam, int lParam, SendMessageTimeoutFlags fuFlags, int uTimeout, out int lpdwResult);
        [DllImport("user32.dll", CharSet = CharSet.Auto, EntryPoint = "SendMessageTimeout", SetLastError = true)]
        public static extern bool SendMessageTimeout(int hWnd, int Msg, int wParam, StringBuilder lParam, SendMessageTimeoutFlags fuFlags, int uTimeout, out int lpdwResult);
        [DllImport("user32.dll")]
        public static extern int FindWindowEx(int hwndParent, int hwndChildAfter, string lpszClass, string lpszWindow);
        [DllImport("user32.dll")]
        public static extern bool IsWindowVisible(int hWnd);
        [DllImport("user32.dll")]
        public static extern int GetWindowRect(int hWnd, ref Rect rect);

        [Flags]
        public enum SendMessageTimeoutFlags : uint
        {
            SMTO_NORMAL = 0u,
            SMTO_BLOCK = 1u,
            SMTO_ABORTIFHUNG = 2u,
            SMTO_NOTIMEOUTIFNOTHUNG = 8u,
            SMTO_ERRORONEXIT = 32u
        }

        public struct Rect
        {
            public int left;
            public int top;
            public int right;
            public int bottom;
        }

        public static Rectangle GetWindowRectangle(int hWnd)
        {
            Rectangle rect;
            if (hWnd > 0)
            {
                Rect defRect = default(Rect);
                GetWindowRect(hWnd, ref defRect);
                rect = new Rectangle(defRect.left, defRect.top, defRect.right - defRect.left, defRect.bottom - defRect.top);
            }
            else
            {
                rect = default(Rectangle);
            }
            return rect;
        }

        #endregion


    }


    public class XYRatio
    {
        public double XRatio;
        public double YRatio;
        public XYRatio()
        {
        }
    }

    public enum ClsNameEnum
    {
        Unknown,
        StandardWindow,
        StandardButton,
        SplitterBar,
        RichEditComponent,
        PrivateWebCtrl,
        Aef_WidgetWin_0,
        ToolBarPlus,
        Aef_RenderWidgetHostHWND,
        SuperTabCtrl,
        StackPanel,
        EditComponent
    }

    public class WindowClue
    {
        public string ClsName
        {
            get;
            private set;
        }
        public string Text
        {
            get;
            private set;
        }
        public int SkipCount
        {
            get;
            private set;
        }
        public WindowClue(ClsNameEnum clsname, string text = null, int skipCount = -1)
        {
            this.ClsName = clsname.ToString();
            if (text == "")
            {
                text = null;
            }
            this.Text = text;
            this.SkipCount = skipCount;
        }
        public WindowClue(string clsname, string text = null, int skipCount = -1)
        {
            this.ClsName = clsname;
            this.Text = ((text == "") ? null : text);
            this.SkipCount = skipCount;
        }
    }
}
