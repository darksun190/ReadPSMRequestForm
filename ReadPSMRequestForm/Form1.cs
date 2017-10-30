using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;
using log4net.Appender;
using System.Resources;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;

namespace ReadPSMRequestForm
{
    public partial class Form1 : Form, log4net.Appender.IAppender
    {
        log4net.ILog logger = log4net.LogManager.GetLogger(typeof(Form1));
        string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "requestlist.xlsx");
        public void DoAppend(log4net.Core.LoggingEvent loggingEvent)
        {
            if (loggingEvent.Level == log4net.Core.Level.Info)
            {
                ReportProgress(String.Format("{0}", loggingEvent.MessageObject.ToString()));
            }
            else
            {
                ReportProgress(String.Format("{0}: {1}", loggingEvent.Level.Name, loggingEvent.MessageObject.ToString()));
            }
        }
        private void ReportProgress(string p)
        {
            StringBuilder sb = new StringBuilder(textBox1.Text);
            if (textBox1.Lines.Length > 100)
            {
                sb.Remove(0, textBox1.Text.IndexOf('\r', 50) + 2);
                sb.Append(p);
                textBox1.Clear();
                textBox1.AppendText(sb.ToString());
            }
            else
            {
                textBox1.AppendText(p);
            }
            textBox1.AppendText(Environment.NewLine);
        }
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
        }
        void convertPDF(string path)
        {
            logger.Info(string.Format("start converting the file {0}", path));
            using (PdfReader reader = new PdfReader(path))
            {
                PSMRequst requst;
                var dic = reader.AcroFields;
                try
                {
                    //prepare for change of pdf, none of using now
                    double version = double.Parse(dic.GetField("Version"));
                    requst = new PSMRequst();
                    requst.parse(dic);

                }
                catch (Exception e)
                {
                    logger.Error("没有读取到相关信息，可能是错误的PDF文件");
                    return;
                }
                AddRequest(requst);
            }
        }

        private void AddRequest(PSMRequst requst)
        {
            XSSFWorkbook workbook;
            string temp_path = path + "." + DateTime.Now.ToOADate().ToString();
            File.Copy(path, temp_path);

            using (var s = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                workbook = new XSSFWorkbook(s);
            }

            XSSFSheet sheet = workbook.GetSheet("Sheet1") as XSSFSheet;

            foreach (var r in sheet.AsIEnumerable())
            {
                if (r.Cells[0].StringCellValue == requst.OpportunityID)
                {
                    logger.Warn("找到了相同的OpportunityID，忽略此文件与信息");
                    return;
                }
            }
            var row = sheet.CreateRow(sheet.LastRowNum + 1);
            AddRequestToRow(row, requst);
            logger.Info(string.Format("已经添加了Opportunity：{0}", requst.OpportunityID));
            try
            {
                using (var s = new FileStream(path, FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(s);
                    logger.Info("文件写入成功");
                }
                File.Delete(temp_path);
            }
            catch (Exception e)
            {
                logger.Error("文件写入失败");
                return;
            }
            finally
            {
                workbook.Close();

            }
        }

        private void AddRequestToRow(IRow row, PSMRequst requst)
        {
            AddCellToRow(row, 0, requst.OpportunityID);
            AddCellToRow(row, 1, requst.SalesName);

            AddCellToRow(row, 2, requst.ApplyDate);
            AddCellToRow(row, 3, requst.PSMName);
            AddCellToRow(row, 4, requst.CustomerName);
            AddCellToRow(row, 5, requst.City);
            AddCellToRow(row, 6, requst.PartType);
            AddCellToRow(row, 7, requst.PaperWork);
            AddCellToRow(row, 8, requst.OnsiteSupport);
            AddCellToRow(row, 9, requst.ProjectManagement);
            AddCellToRow(row, 10, requst.MachineType);
            AddCellToRow(row, 11, requst.MachineNumber);
            AddCellToRow(row, 12, requst.Budget.ToString());
            AddCellToRow(row, 13, requst.CompetitorInfo);
            AddCellToRow(row, 14, requst.ExistingZeissInfo);
            AddCellToRow(row, 15, requst.ExistingCompetitorInfo);
            AddCellToRow(row, 16, requst.AdvantageAndDisadvantage);
            AddCellToRow(row, 17, requst.Remarks);
        }

        private void AddCellToRow(IRow row, int v, bool b)
        {
            var c = row.CreateCell(v);
            if (b)
                c.SetCellValue("Yes");
            else
                c.SetCellValue("No");
        }

        private void AddCellToRow(IRow row, int v, string s)
        {
            var c = row.CreateCell(v);
            c.SetCellValue(s);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Multiselect = true;
                ofd.RestoreDirectory = true;
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    checkSpreadSheet();
                    foreach (var f in ofd.FileNames)
                    {
                        if (f.EndsWith(".pdf"))
                        {
                            convertPDF(f);
                        }
                        else
                        {
                            logger.Warn(string.Format("{0} isn't a pdf file, ignore it", f));
                        }
                    }
                }
            }
        }

        private void checkSpreadSheet()
        {
            if (!File.Exists(path))
            {
                var fs = File.Create(path);
                var f = Properties.Resources.requestlist;
                fs.Write(f, 0, f.Count());
                fs.Close();
            }
        }
    }
}
