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
    public interface IPaymentService
    {
        bool SavePayment(IList<Payment> payments, IList<PaymentAllocation> paymentAllocations);

        IList<Entity.Booking.Fee> GetPaymentTypeFee(string originRcd,
                                        string destinationRcd,
                                        string formOfPayment,
                                        string formOfPaymentSubtype,
                                        string currencyRcd,
                                        string agency,
                                        DateTime feeDate);


    }
}
