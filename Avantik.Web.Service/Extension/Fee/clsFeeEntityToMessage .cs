using Avantik.Web.Service.Message.Agency;
using Avantik.Web.Service.Message.Fee;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Avantik.Web.Service.Extension
{
    public static class FeeEntityToMessage 
    {
        public static IList<Message.Fee.Fee> ToFeeMessage(this IList<Entity.Fee> objEntitytFeeList)
        {
            List<Message.Fee.Fee> objMessageFeeList = null;
            if (objEntitytFeeList != null)
            {
                objMessageFeeList = new List<Message.Fee.Fee>();
                for (int i = 0; i < objEntitytFeeList.Count; i++)
                {
                    objMessageFeeList.Add(objEntitytFeeList[i].ToFeeMessage());
                }
            }
            return objMessageFeeList;
        }

        public static Message.Fee.Fee ToFeeMessage(this Entity.Fee f)
        {
            Message.Fee.Fee fee = null;

            if(f != null)
            {
                fee = new Message.Fee.Fee();
                fee.AccountFeeBy = f.AccountFeeBy;
                fee.AccountFeeDateTime = f.AccountFeeDateTime;
                fee.AcctFeeAmount = f.AcctFeeAmount;
                fee.AcctFeeAmountIncl = f.AcctFeeAmountIncl;
                fee.AgencyCode = f.AgencyCode;
                fee.BaggageFeeOptionId = f.BaggageFeeOptionId;
                fee.BookingFeeId = f.BookingFeeId;
                fee.BookingId = f.BookingId;
                fee.BookingSegmentId = f.BookingSegmentId;
                fee.ChangeComment = f.ChangeComment;
                fee.ChargeAmount = f.ChargeAmount;
                fee.ChargeAmountIncl = f.ChargeAmountIncl;
                fee.ChargeCurrencyRcd = f.ChargeCurrencyRcd;
                fee.Comment = f.Comment;
                fee.CreateBy = f.CreateBy;
                fee.CreateDateTime = f.CreateDateTime;
                fee.CurrencyRcd = f.CurrencyRcd;
                fee.DestinationRcd = f.DestinationRcd;
                fee.DisplayName = f.DisplayName;
                fee.DocumentDateTime = f.DocumentDateTime;
                fee.DocumentNumber = f.DocumentNumber;
                fee.ExternalReference = f.ExternalReference;
                fee.FeeAmount = f.FeeAmount;
                fee.FeeAmountIncl = f.FeeAmountIncl;
                fee.FeeCalculationRcd = f.FeeCalculationRcd;
                fee.FeeCategoryRcd = f.FeeCategoryRcd;
                fee.FeeId = f.FeeId;
                fee.FeeLevel = f.FeeLevel;
                fee.FeePercentage = f.FeePercentage;
                fee.FeeRcd = f.FeeRcd;
                fee.ManualFeeFlag = f.ManualFeeFlag;
                fee.MinimumFeeAmountFlag = f.MinimumFeeAmountFlag;
                fee.NewRecord = f.NewRecord;
                fee.NumberOfUnits = f.NumberOfUnits;
                fee.OdDestinationRcd = f.OdDestinationRcd;
                fee.OdFlag = f.OdFlag;
                fee.OdOriginRcd = f.OdOriginRcd;
                fee.OriginRcd = f.OriginRcd;
                fee.PassengerId = f.PassengerId;
                fee.PassengerSegmentServiceId = f.PassengerSegmentServiceId;
                fee.PaymentAmount = f.PaymentAmount;
                fee.SelectedFee = f.SelectedFee;
                fee.SkipFareAllowanceFlag = f.SkipFareAllowanceFlag;
                fee.TotalAmount = f.TotalAmount;
                fee.TotalAmountIncl = f.TotalAmountIncl;
                fee.TotalFeeAmount = f.TotalFeeAmount;
                fee.TotalFeeAmountIncl = f.TotalFeeAmountIncl;
                fee.Units = f.units;
                fee.UpdateBy = f.UpdateBy;
                fee.UpdateDateTime = f.UpdateDateTime;
                fee.VatPercentage = f.VatPercentage;
                fee.VendorRcd = f.VendorRcd;
                fee.VoidBy = f.VoidBy;
                fee.VoidDateTime = f.VoidDateTime;
            }

            return fee;
        }

        public static IList<Message.Fee.ServiceFee> ToFeeMessage(this IList<Entity.ServiceFee> objEntitytFeeList)
        {
            List<Message.Fee.ServiceFee> objMessageFeeList = null;
            if (objEntitytFeeList != null)
            {
                objMessageFeeList = new List<Message.Fee.ServiceFee>();
                for (int i = 0; i < objEntitytFeeList.Count; i++)
                {
                    objMessageFeeList.Add(objEntitytFeeList[i].ToFeeMessage());
                }
            }
            return objMessageFeeList;
        }

        public static Message.Fee.ServiceFee ToFeeMessage(this Entity.ServiceFee f)
        {
            Message.Fee.ServiceFee fee = null;

            if (f != null)
            {
                fee = new Message.Fee.ServiceFee();
                fee.FeeRcd = f.FeeRcd;
                fee.DisplayName = f.DisplayName;
                fee.CurrencyRcd = f.CurrencyRcd;
                fee.FeeAmount = f.FeeAmount;
                fee.FeeAmountIncl = f.FeeAmountIncl;
                fee.TotalFeeAmount = f.TotalFeeAmount;
                fee.TotalFeeAmountIncl = f.TotalFeeAmountIncl;
                fee.ServiceOnRequestFlag = f.ServiceOnRequestFlag;
                fee.CutOffTime = f.CutOffTime;

                // part of SegmentService
                fee.FlightConnectionId = f.FlightConnectionId;
                fee.SpecialServiceRcd = f.SpecialServiceRcd;
                fee.OriginRcd = f.OriginRcd;
                fee.DestinationRcd = f.DestinationRcd;
                fee.OdOriginRcd = f.OdOriginRcd;
                fee.OdDestinationRcd = f.OdDestinationRcd;
                fee.BookingClassRcd = f.BookingClassRcd;
                fee.FareCode = f.FareCode;
                fee.AirlineRcd = f.AirlineRcd;
                fee.FlightNumber = f.FlightNumber;
                fee.DepartureDate = f.DepartureDate;

            }

            return fee;
        }


    }
}
