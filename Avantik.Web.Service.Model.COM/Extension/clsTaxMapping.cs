using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Avantik.Web.Service.Entity;
using Avantik.Web.Service.COMHelper;
using Avantik.Web.Service.Exception.Booking;

namespace Avantik.Web.Service.Model.COM.Extension
{
    public static class TaxMapping
    {
        // add tax to mapping
        public static IList<Avantik.Web.Service.Entity.Booking.Mapping> FillTaxMapping(this  IList<Avantik.Web.Service.Entity.Booking.Tax> taxList, IList<Avantik.Web.Service.Entity.Booking.Mapping> m)
        {

            if (taxList != null && m != null)
            {
                for (int i = 0; i < taxList.Count; i++)
                {
                    for (int j = 0; j < m.Count; j++)
                    {
                        if (taxList[i].BookingSegmentId == m[j].BookingSegmentId && taxList[i].PassengerId == m[j].PassengerId)
                            taxList[i].FillTaxMapping(m[j]);
                    }
                }
            }
            return m;
        }

        public static Avantik.Web.Service.Entity.Booking.Mapping FillTaxMapping(this  Avantik.Web.Service.Entity.Booking.Tax tax, Avantik.Web.Service.Entity.Booking.Mapping mapping)
        {
            if (tax != null && mapping != null)
            {

                if (tax.TaxRcd != null && tax.TaxRcd.ToUpper() == "YQ")
                {
                    mapping.YqAmount += tax.TaxAmount;
                    mapping.YqAmountIncl += tax.TaxAmountIncl;

                    mapping.AcctYqAmount += tax.TaxAmount;
                    mapping.AcctYqAmountIncl += tax.TaxAmountIncl;

                    mapping.YqVat += tax.TaxAmountIncl - tax.TaxAmount;
                }
                else
                {
                    mapping.TaxAmount += tax.TaxAmount;
                    mapping.TaxAmountIncl += tax.TaxAmountIncl;
                    mapping.TaxVat += tax.TaxAmountIncl - tax.TaxAmount;

                    mapping.AcctTaxAmount += tax.TaxAmount;
                    mapping.AcctTaxAmountIncl += tax.TaxAmountIncl;
                }

            }

            return mapping;
        }

    }
}
