using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace Avantik.Web.Service.Entity.Route
{
    public class Route
    {
        public string OriginRcd { get; set; }
        public string DestinationRcd { get; set; }
        public bool SegmentChangeFeeFlag { get; set; }
        public bool TransitFlag { get; set; }
        public bool DirectFlag { get; set; }
        public bool AvlFlag { get; set; }
        public bool B2CFlag { get; set; }
        public bool B2BFlag { get; set; }
        public bool B2TFlag { get; set; }
        public Int16 DayRange { get; set; }
        public bool ShowRedressNumberFlag { get; set; }
        public bool RequirePassengerTitleFlag { get; set; }
        public bool RequirePassengerGenderFlag { get; set; }
        public bool RequireDateOfBirthFlag { get; set; }
        public bool RequireDocumentDetailsFlag { get; set; }
        public bool RequirePassengerWeightFlag { get; set; }
        public bool SpecialServiceFeeFlag { get; set; }
        public bool ShowInsuranceOnWebFlag { get; set; }
        public string DisplayName { get; set; }
        public string CountryRcd { get; set; }
        public string CurrencyRcd { get; set; }
        public byte RoutesTot { get; set; }
        public byte RoutesAvl { get; set; }
        public byte RoutesB2C { get; set; }
        public byte RoutesB2B { get; set; }
        public byte RoutesB2S { get; set; }
        public byte RoutesAPI { get; set; }
        public byte RoutesB2T { get; set; }

    }
}
