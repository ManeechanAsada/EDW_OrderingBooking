using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Avantik.Web.Service.Entity.Flight;
using Avantik.Web.Service.Infrastructure;
using Avantik.Web.Service.Entity.Booking;
using Avantik.Web.Service.Entity.Agency;
using Avantik.Web.Service.Entity;

namespace Avantik.Web.Service.Model.Contract
{
    public interface IVoucherService
    {
        Voucher GetVoucher(Voucher Voucher, IList<Entity.Booking.Mapping> Mappings, IList<Entity.Booking.Fee> Fees, string agencyCode, string userId);   
    }
}
