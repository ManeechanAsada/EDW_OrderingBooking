using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.ServiceModel;

namespace Avantik.Web.Service.Message.OrderBooking
{
    [MessageContract]
    public class Payment
    {
        //If not passin use server time.
        [MessageBodyMember]
        public DateTime payment_date_time { get; set; }
        
        //Electronic fund transfer information.
       // public EFT EFT { get; set; }
        
        //Credit card information.
       // public CreditCard credit_card { get; set; }
        [MessageBodyMember]
        public decimal payment_amount { get; set; }
        [MessageBodyMember]
        public decimal receive_payment_amount { get; set; }

        //Value will be SALE or REFUND
        [MessageBodyMember]
        public string payment_type_rcd { get; set; }
        [MessageBodyMember]
        public string form_of_payment_rcd { get; set; }
        [MessageBodyMember]
        public string form_of_payment_subtype_rcd { get; set; }
        [MessageBodyMember]
        public string debit_agency_code { get; set; }
        [MessageBodyMember]
        public string currency_rcd { get; set; }
        [MessageBodyMember]
        public string receive_currency_rcd { get; set; }
        [MessageBodyMember]
        public string approval_code { get; set; }
        [MessageBodyMember]
        public string transaction_reference { get; set; }
        [MessageBodyMember]
        public string payment_number { get; set; }
        [MessageBodyMember]
        public string payment_reference { get; set; }
        [MessageBodyMember]
        public string payment_remark { get; set; }
    }
}
