using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace 获取千牛聊天对象
{
    public class ChatRecordParser
    {
        protected string _User = string.Empty;
        protected Dictionary<string, ChatRecord> ChatRecordCache = new Dictionary<string, ChatRecord>(30, null);
        protected Dictionary<string, CBStatement> _StatementCache = new Dictionary<string, CBStatement>(500, null);
        public ChatRecordParser(string user)
        {
            this._User = user;
        }
        public ChatRecord Parse(string html)
        {
            ChatRecord chatRecord = this.ChatRecordCache.ContainsKey(html) ? this.ChatRecordCache[html] : null;
            if (chatRecord == null)
            {
                var cbs = this.ParseStatements(html);
                if (cbs != null)
                {
                    chatRecord = new ChatRecord(cbs, this._User);
                    this.ChatRecordCache.Add(html, chatRecord);
                }
            }
            return chatRecord;
        }
        protected List<CBStatement> ParseStatements(string html)
        {
            List<CBStatement> cbs = null;
            if (string.IsNullOrEmpty(html)) return cbs;

            HtmlDocument htmlDocument = new HtmlDocument();
            htmlDocument.LoadHtml(html);
            cbs = this.ParseStatementsInner(htmlDocument);
            return cbs;
        }
        protected virtual List<CBStatement> ParseStatementsInner(HtmlDocument htmlDoc)
        {
            List<CBStatement> cbsLst = new List<CBStatement>();
            HtmlNode elementbyId = htmlDoc.GetElementbyId("J_msgContainer");
            HtmlNodeCollection htmlNodes = (elementbyId != null) ? elementbyId.SelectNodes(".//div[contains(@class,'J_msg imui-msg ')]") : null;
            if (htmlNodes != null)
            {
                foreach (HtmlNode node in htmlNodes)
                {
                    try
                    {
                        if (!node.Attributes["class"].Value.Contains("imui-msg-system"))
                        {
                            string outerHtml = node.OuterHtml;
                            CBStatement cBStatement;
                            if (this._StatementCache.ContainsKey(outerHtml))
                            {
                                cBStatement = this._StatementCache[outerHtml];
                            }
                            else
                            {
                                cBStatement = this.CreateCBStatement(node);
                                if (cBStatement != null)
                                {
                                    this._StatementCache.Add(outerHtml, cBStatement);
                                }
                            }
                            if (cBStatement != null)
                            {
                                cbsLst.Add(cBStatement);
                            }
                        }
                    }
                    catch (OutOfMemoryException ex)
                    {
                        GC.Collect();
                    }
                    catch (Exception ex2)
                    {
                    }
                }
            }
            return cbsLst;
        }
        private CBStatement CreateCBStatement(HtmlNode htmlNode)
        {
            CBStatement cbs = null;
            for (int i = 0; i < 2; i++)
            {
                try
                {
                    string value = htmlNode.Attributes["data-time"].Value;
                    DateTime time = ChatRecordParser.TimeSpanToDateTime(Convert.ToInt64(value));
                    string nick = htmlNode.Attributes["data-nick"].Value;
                    bool isSellerMsg = htmlNode.Attributes["class"].Value.Contains("imui-msg-r");
                    htmlNode.Attributes["class"].Value.Contains("imui-msg-l");
                    CBStatement.Role role;
                    if (isSellerMsg)
                    {
                        role = CBStatement.Role.Self;
                    }
                    else
                    {
                        role = CBStatement.Role.Guest;
                    }
                    HtmlNode htmlNode_ = htmlNode.SelectSingleNode(".//div[contains(@class,'msg-body-html')]");
                    string chatmsg = this.GetChatContent(htmlNode_);
                    cbs = new CBStatement(nick, chatmsg, time, role);
                    return cbs;
                }
                catch (OutOfMemoryException ex)
                {
                    GC.Collect();
                }
                catch (Exception ex2)
                {
                }
            }
            return cbs;
        }
        private string GetChatContent(HtmlNode htmlNode)
        {
            string text = this.GetText(htmlNode.ChildNodes);
            text = text.Replace("&nbsp;", " ");
            text = text.Replace("&gt;", ">");
            text = text.Replace("&lt;", "<");
            return text.Replace("&amp;", "&");
        }
        private string GetText(HtmlNodeCollection htmlNodes)
        {
            StringBuilder sb = new StringBuilder();
            if (htmlNodes == null || htmlNodes.Count < 1) return string.Empty;

            foreach (HtmlNode node in htmlNodes)
            {
                if (!(node.Name == "div"))
                {
                    if (node.Name == "#text")
                    {
                        sb.Append(node.InnerText);
                    }
                    else
                    {
                        if (node.Name == "img")
                        {
                            sb.Append(node.InnerText);
                        }
                        if (node.ChildNodes != null && node.ChildNodes.Count > 0)
                        {
                            sb.Append(this.GetText(node.ChildNodes));
                        }
                    }
                }

            }
            return sb.ToString().Trim();
        }
        public static DateTime TimeSpanToDateTime(long timeSpan)
        {
            DateTime result = new DateTime(1970, 1, 1, 0, 0, 0);
            result = result.AddMilliseconds((double)timeSpan).ToLocalTime();
            return result;
        }

    }

}
