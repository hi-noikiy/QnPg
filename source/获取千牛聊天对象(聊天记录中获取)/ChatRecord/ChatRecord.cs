using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace 获取千牛聊天对象
{
    public class ChatRecord
    {
        public List<CBStatement> Statements;
        private string _Seller = string.Empty;
        private string Speaker = string.Empty;
        public string GuestName
        {
            get
            {
                return this.Speaker;
            }
            private set
            {
                this.Speaker = value;
            }
        }
        public ChatRecord(List<CBStatement> cbslist, string seller)
        {
            this.Statements = cbslist;
            this._Seller = seller;
            if (cbslist != null && cbslist.Count > 0)
            {
                for (int i = cbslist.Count - 1; i >= 0; i--)
                {
                    CBStatement cBStatement = cbslist[i];
                    if (cBStatement.SpeakerRole == CBStatement.Role.Guest)
                    {
                        this.Speaker = cBStatement.Speaker;
                        break;
                    }
                }
            }
        }
        public bool HasStockScreen()
        {
            bool isStockScreen = false;
            if (this.Statements != null && this.Statements.Count > 0)
            {
                int index = this.Statements.Count - 1;
                isStockScreen = (this.Statements[index].SpeakerRole == CBStatement.Role.SysMsg
                    && this.Statements[index].OneStringWithTrim.Contains("振屏"));
            }
            return isStockScreen;
        }
    }

}
