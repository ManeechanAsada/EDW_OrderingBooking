using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Avantik.Web.Service.Entity
{
    public  class Authentication 
    {
        public string Token { get; set; }
        public Guid UserId { get; set; }
        public string AgencyCode { get; set; }
        public string B2bAllowSeat { get; set; }
        public string B2bAllowService { get; set; }
        public string B2bAllowInFoPassenger { get; set; }
        public string B2bAllowFlightChange { get; set; }
        public string B2bCancelFlight { get; set; }
        public string B2bAllowNameChange{ get; set; }
        public string B2bAllowChangeDetail { get; set; }

        public string CurrencyRcd { get; set; }
        public string LanguageRcd { get; set; }

        public bool ResponseSuccess { get; set; }
        public string ResponseMessage { get; set; }
        public string ResponseCode { get; set; }

    }
}
