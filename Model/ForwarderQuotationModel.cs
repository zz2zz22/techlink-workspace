using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace techlink_workspace.Model
{
    public class ForwarderQuotationModel
    {
        public string Forwarder_ID { get; set; }
        public string Forwarder_name { get; set; }
        public string Forwarder_portDelivery { get; set; }
        public string Forwarder_term { get; set; }
        public string Forwarder_container { get; set; }
        public string Forwarder_commodity { get; set; }
        public string Forwarder_hsCode { get; set; }
        public string Forwarder_carrier { get; set; }
        public decimal? Forwarder_total { get; set; }
        public decimal? Forwarder_of { get; set; }
        public decimal? Forwarder_localPol { get; set; }
        public decimal? Forwarder_destCharge { get; set; }
        public decimal? Forwarder_delivery { get; set; }
        public decimal? Forwarder_otherCharge { get; set; }
        public string Forwarder_remark { get; set; }
        public int? Forwarder_volumn { get; set; }
        public string Forwarder_validDate { get; set; }
        public DateTime? create_date { get; set; }
        public string create_by { get; set; }
        public DateTime? update_date { get; set; }
        public string update_by { get; set; }
    }
}
