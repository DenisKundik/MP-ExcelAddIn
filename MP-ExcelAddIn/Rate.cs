using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MP_ExcelAddIn
{
    public class Rate
    {
        public String ExecutorName { get; set; }
        public Double SalesSupportRate { get; set; }
        public Double ProjectRate { get; set; }
        public Double ProjectOvertimeRate { get; set; }
        public Double CustomerRate { get; set; }
        public Double CustomerOvertimeRate { get; set; }

        public Rate (string executorName, double salesSupportRate, double projectRate, double projectOvertimeRate, double customerRate, double customerOvertimeRate)
        {
            ExecutorName = executorName;
            SalesSupportRate = salesSupportRate;
            ProjectRate = projectRate;
            ProjectOvertimeRate = projectOvertimeRate;
            CustomerRate = customerRate;
            CustomerOvertimeRate = customerOvertimeRate;
        }
    }
}
