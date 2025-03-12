using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Avantik.Web.Service.Message
{
    public class RouteView
    {
        public string origin_rcd { get; set; }
        public string destination_rcd { get; set; }
        public bool segment_change_fee_flag { get; set; }
        public bool transit_flag { get; set; }
        public bool direct_flag { get; set; }
        public bool avl_flag { get; set; }
        public bool b2c_flag { get; set; }
        public bool b2b_flag { get; set; }
        public bool b2t_flag { get; set; }
        public bool b2e_flag { get; set; }
        public bool b2s_flag { get; set; }
        public bool api_flag { get; set; }
        public Int16 day_range { get; set; }
        public bool show_redress_number_flag { get; set; }
        public bool require_passenger_title_flag { get; set; }
        public bool require_passenger_gender_flag { get; set; }
        public bool require_date_of_birth_flag { get; set; }
        public bool require_document_details_flag { get; set; }
        public bool require_passenger_weight_flag { get; set; }
        public bool show_special_service_on_web_flag { get; set; }
        public bool show_insurance_on_web_flag { get; set; }
        public string display_name { get; set; }
        public string country_rcd { get; set; }
        public string currency_rcd { get; set; }
        public byte routes_tot { get; set; }
        public byte routes_avl { get; set; }
        public byte routes_b2c { get; set; }
        public byte routes_b2b { get; set; }
        public byte routes_b2s { get; set; }
        public byte routes_api { get; set; }
        public byte routes_b2t { get; set; }
        public string language { get; set; }
    }
}
