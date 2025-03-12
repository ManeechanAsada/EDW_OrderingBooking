using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.ServiceModel;
using System.Runtime.Serialization;
using Avantik.Web.Service;
using Avantik.Web.Service.Message.Fee;

namespace Avantik.Web.Service.Message.ManageBooking
{
    [MessageContract]
    public class ModifyBookingRequest
    {
        [MessageHeader]
        public string Token { get; set; }
        [MessageBodyMember]
        public string BookingId { set; get; }
        [MessageBodyMember]
        public IList<ChangeFlight> ModifyFlights { get; set; }
        [MessageBodyMember]
        public IList<SSR.Service> ModifyPassengerServices { set; get; }
        [MessageBodyMember]
        public IList<SeatAssign> ModifySeats { set; get; }
        [MessageBodyMember]
        public IList<BaggageRequest> ModifyBaggages { set; get; }
        [MessageBodyMember]
        public IList<NameChange> ModifyPassengerName { set; get; }
        [MessageBodyMember]
        public IList<Payment> Payments { get; set; }
        [MessageBodyMember]
        public string ActionCode { set; get; }
        [MessageBodyMember]
        public bool SkipPaymentFeeFlag { set; get; }

        [MessageBodyMember]
        public IList<Message.ManageBooking.Fee> Fees { get; set; }

    }
}
