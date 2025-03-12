using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace Avantik.Web.Service.Entity.FormOfPayment
{
    public class FormOfPayment 
    {
        #region Property

        public IList<FormOfPaymentSubType> FormOfPaymentSubType { get; set; }

        #endregion
    }
}
