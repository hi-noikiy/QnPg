---
author: 小飞虾
comments: true
date: 2018-09-27 13:11
layout: post
title: 千牛聊天记录获取
categories:
- 千牛插件开发
---

## 分析UI:
分析千牛UI控件，我们用Visual Studio自带的SPY++查找窗口，得到聊天记录的控件信息发现 **窗口类名:Aef_RenderWidgetHostHWND** ,上网搜了一下说是Chrominum 的窗口。确定一下我们直接选中千牛的聊天窗口按**F12**,发现会弹出Chrome的开发者工具。到此我们确定了千牛的聊天窗口就是内嵌了一个Chrome
## 思路：
[利用Chrome远程调试技术](http://taobaofed.org/blog/2016/10/19/chrome-remote-debugging-technics/ "揭秘浏览器远程调试技术"),获取Chrome的页面的信息
下面是我本地千牛获取到的千牛Debug的信息
[![qianiudebug](http://7xpf2l.com1.z0.glb.clouddn.com/qianniudebug.gif "qianiudebug")](http://7xpf2l.com1.z0.glb.clouddn.com/qianniudebug.gif "qianiudebug")
我们可以看到可以获取到聊天窗口。
## 编码实现
需要用到一个开源库，[ChromeDevTools](https://github.com/MasterDevs/ChromeDevTools "ChromeDevTools")

1、先用WINAPI获取千牛的工作台窗口句柄(不是聊天记录窗口的句柄)。
```C#
 public virtual Dictionary<string, int> GetAllChatDeskSellerNameAndHwndInner()
        {
            Dictionary<string, int> rtdict = new Dictionary<string, int>();
            FindAllDesktopWindowByClassNameAndTitlePattern("StandardFrame", chatWindowTitlePattern, (int hwnd, string title) =>
            {
                if (IsWindowVisible(hwnd))
                {
                    string winTitle = Regex.Match(title, chatWindowTitlePattern).ToString();
                    if (!string.IsNullOrEmpty(winTitle)){
                        rtdict[winTitle] = hwnd;
                    }
                }
            });
            return rtdict;
        }
```
2、获取Chrome远程调试端口
```C#
//获取进程id
int pid = GetWindowThreadProcessId(hwnd);
//获取chrome调试端口
int port = GetChromeListeningPortWithInterOp(pid);

var sessionInfoUrl =  string.Format("http://localhost:{0}", port);
```
3、获取千牛聊天窗口内的页面webSocketDebuggerUrl ，获取聊天记录
```C#
//获取所有的DebugSessionInfo
var sessionInfos = GetWebSocketSessionInfos(sessionInfoUrl);
//获取聊天记录窗口DebugSessionInfo
var sessionInfo = sessionInfos = sessionInfos.First(k =>k.Title.Contains("聊天窗口")).ToList();
//CreateChromeSession
var chromeSession = new ChromeSessionFactory().Create(webSocketSessionInfo.WebSocketDebuggerUrl); 
//获取聊天记录
var html = GetHtml(out html);

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

```

##以下是最终实现效果
[![qnchatlog](http://7xpf2l.com1.z0.glb.clouddn.com/qianniugetchatlog.gif "qnchatlog")](http://7xpf2l.com1.z0.glb.clouddn.com/qianniugetchatlog.gif "qnchatlog")










