using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using iTextSharp.text.pdf;

namespace ReadPSMRequestForm
{
    public class PSMRequst
    {
        public string SalesName
        {
            get;
            set;
        }
        public string ApplyDate
        {
            get;
            set;
        }
        public string PSMName
        {
            get;
            set;
        }
        public string CustomerName
        {
            get;
            set;
        }
        public string City
        {
            get;
            set;
        }
        public bool NewProject
        {
            get;
            set;
        }
        public bool ExtendProject
        {
            get;
            set;
        }
        public bool RFQFixed
        {
            get;
            set;
        }
        public string PartType
        {
            get;
            set;
        }
        public bool PaperWork
        {
            get;
            set;
        }
        public bool OnsiteSupport
        {
            get;
            set;
        }
        public bool ProjectManagement
        {
            get;
            set;
        }
        public string OpportunityID
        {
            get;
            set;
        }
        public string MachineType
        {
            get;
            set;
        }
        public string MachineNumber
        {
            get;
            set;
        }
        public double Budget
        {
            get;
            set;
        }
        public string CompetitorInfo
        {
            get;
            set;
        }
        public string ExistingZeissInfo
        {
            get;
            set;
        }
        public string ExistingCompetitorInfo
        {
            get;
            set;
        }
        public string AdvantageAndDisadvantage
        {
            get;
            set;
        }
        public string Remarks
        {
            get;
            set;
        }

        internal void parse(AcroFields dic)
        {
            this.OpportunityID = dic.StringFieldValue("OpportunityID");
            this.SalesName = dic.StringFieldValue("SalesName");
            this.ApplyDate = dic.StringFieldValue("ApplyDate");
            this.PSMName = dic.StringFieldValue("PSMName");
            this.CustomerName = dic.StringFieldValue("CustomerName");
            this.City = dic.StringFieldValue("City");
            this.NewProject = dic.BoolFieldValue("CheckNewProject");
            this.ExtendProject = dic.BoolFieldValue("CheckExtendProject");
            this.RFQFixed = dic.BoolFieldValue("CheckRFQFixed");
            this.PartType = dic.StringFieldValue("PartType");
            this.PaperWork = dic.BoolFieldValue("CheckPaperWork");
            this.OnsiteSupport = dic.BoolFieldValue("CheckOnsiteSupport");
            this.ProjectManagement = dic.BoolFieldValue("CheckProjectManagement");
            this.MachineType = dic.StringFieldValue("MachineType");
            this.MachineNumber = dic.StringFieldValue("MachineNumber");
            this.Budget = double.Parse(dic.StringFieldValue("Budget"));
            this.CompetitorInfo = dic.StringFieldValue("CompetitorInfo");
            this.ExistingZeissInfo = dic.StringFieldValue("ExistingZeissInfo");
            this.ExistingCompetitorInfo = dic.StringFieldValue("ExistingCompetitorInfo");
            this.AdvantageAndDisadvantage = dic.StringFieldValue("AdvantageAndDisadvantage");
            this.Remarks = dic.StringFieldValue("Remarks");
          
        }
    }
}
