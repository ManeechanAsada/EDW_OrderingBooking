using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
namespace Avantik.Web.Service.Entity
{
    public class SeatAssign
    {
        public string BookingSegmentID { get; set; }
        public string PassengerID { get; set; }
        public Int32 SeatRow { get; set; }
        public string SeatColumn { get; set; }
        public string SeatNumber { get; set; }
        public string SeatFeeRcd { get; set; }
    }
}
