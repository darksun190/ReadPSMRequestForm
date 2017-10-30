using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using iTextSharp.text.pdf;

namespace ReadPSMRequestForm
{
    public static class ExtendAcroFields
    {
        public static string StringFieldValue(this AcroFields Dic,string keyname)
        {
            if(Dic.Fields.ContainsKey(keyname))
            {
                return Dic.GetField(keyname);
            }
            else
            {
                return "";
            }
        }
        public static bool BoolFieldValue(this AcroFields Dic, string keyname)
        {
            if (Dic.Fields.ContainsKey(keyname))
            {
                var v = Dic.GetField(keyname);
                if(string.IsNullOrEmpty(v))
                {
                    return false;
                }
                else if (v.ToLower() == "yes" || v == "是")
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
        }
    }
}
