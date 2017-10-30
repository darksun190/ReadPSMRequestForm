using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.XSSF.UserModel;

namespace ReadPSMRequestForm
{
    static class ExtendNPOI
    {
        /// <summary>
        /// 将Sheet表转化为IEnumerable
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns></returns>
        public static IEnumerable<XSSFRow> AsIEnumerable(this XSSFSheet sheet)
        {
            var itor = sheet.GetRowEnumerator();
            while (itor.MoveNext())
            {
                yield return (XSSFRow)itor.Current;
            }
        }
        /// <summary>
        /// 将Excel的行转化为IEnumerable
        /// </summary>
        /// <param name="row"></param>
        /// <returns></returns>
        public static IEnumerable<XSSFCell> AsIEnumerable(this XSSFRow row)
        {
            var itor = row.GetEnumerator();
            while (itor.MoveNext())
            {
                yield return (XSSFCell)itor.Current;
            }
        }
    }
}
