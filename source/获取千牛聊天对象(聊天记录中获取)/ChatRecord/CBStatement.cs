using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace 获取千牛聊天对象
{
    public class CBStatement
    {
        public enum Role
        {
            Self,
            Guest,
            SysMsg,
            Unknown
        }
        public List<string> Lines = new List<string>();
        public DateTime Time;
        public string goodsname;
        public string Speaker;
        public Role SpeakerRole = Role.Unknown;
        public string OneStringWithTrim
        {
            get
            {
                string text = string.Empty;
                foreach (string line in this.Lines)
                {
                    text = text + line + Environment.NewLine;
                }
                return text.Trim();
            }
        }
        public CBStatement(string speaker, string text, DateTime time, Role role)
        {
            this.Lines = new List<string>
			{
				text
			};
            this.Time = time;
            this.Speaker = speaker;
            this.SpeakerRole = role;
        }
        public CBStatement()
        {
        }
        public static CBStatement DeeplyCopy(CBStatement cbstatement)
        {
            CBStatement cBStatement = new CBStatement();
            cBStatement.Time = cbstatement.Time;
            cBStatement.goodsname = cbstatement.goodsname;
            cBStatement.Speaker = cbstatement.Speaker;
            for (int i = 0; i < cbstatement.Lines.Count; i++)
            {
                cBStatement.Lines.Add(cbstatement.Lines[i]);
            }
            return cBStatement;
        }
    }

}
