using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using Avantik.Web.Service.Entity.Flight;
using Avantik.Web.Service.Repository.Contract.Flight;

namespace Avantik.Web.Service.Model
{
    public class AvailabilityTax : AvailabilityDecorator
    {
        protected IFlightRepository _flightRepository;

        string _originRcd;
        string _destinationRcd;
        string _otherPaxType;
        string _currencyRcd;
        string _agencyCode;

        byte _adult;
        byte _child;
        byte _infant;
        
        DateTime _fromDate;
        DateTime _toDate;
        DateTime _bookingDate;

        bool _noVat;
        public AvailabilityTax(IAvailabilityBase availability, 
                                IFlightRepository flightRepository,
                                string originRcd,
                                string destinationRcd,
                                string otherPaxType,
                                string currencyRcd,
                                string agencyCode,
                                DateTime fromDate,
                                DateTime toDate,
                                DateTime bookingDate,
                                byte adult,
                                byte child,
                                byte infant,
                                bool noVat) 
            : base(availability)
        {
            _flightRepository = flightRepository;

            _originRcd = originRcd;
            _destinationRcd = destinationRcd;
            _currencyRcd = currencyRcd;
            _agencyCode = agencyCode;

            _fromDate = fromDate;
            _toDate = toDate;
            _bookingDate = bookingDate;

            _adult = adult;
            _child = child;
            _infant = infant;
            _otherPaxType = otherPaxType;

            _noVat = noVat;
        }

        public override IList<Availability> GetAvailability()
        {
            try
            {
                IList<Availability> baseAvailability = base._Availability.GetAvailability();
                if (baseAvailability == null)
                {
                    throw new ArgumentNullException("availability", "Constructure parameter can not be null.");
                }
                else if (_flightRepository == null)
                {
                    throw new ArgumentNullException("flightRepository", "Constructure parameter can not be null.");
                }
                else
                {
                    //Find available tax for calculate.
                    
                    if (baseAvailability.FindAvlNetTotalFlag() == true)
                    {
                        //Loop through date range for each pax type.
                        IList<AvailabilityQuoteTax> availabilityQuoteTax = new List<AvailabilityQuoteTax>();
                        IEnumerable<AvailabilityQuoteTax> tempAvailabilityQuoteTax = null;
                        DateTime dtCurrent = _fromDate;

                        while (dtCurrent.Date <= _toDate.Date)
                        {
                            if (_adult > 0)
                            {
                                tempAvailabilityQuoteTax = _flightRepository.GetFlightAvailabilityQuoteTax(_originRcd,
                                                                                                           _destinationRcd,
                                                                                                           _currencyRcd,
                                                                                                           "ADULT",
                                                                                                           _agencyCode,
                                                                                                           dtCurrent,
                                                                                                           _bookingDate);

                                FillAvaiTax(availabilityQuoteTax, tempAvailabilityQuoteTax);
                                tempAvailabilityQuoteTax = null;
                            }
                            if (_child > 0)
                            {
                                tempAvailabilityQuoteTax = _flightRepository.GetFlightAvailabilityQuoteTax(_originRcd,
                                                                                                           _destinationRcd,
                                                                                                           _currencyRcd,
                                                                                                           "CHD",
                                                                                                           _agencyCode,
                                                                                                           dtCurrent,
                                                                                                           _bookingDate);

                                FillAvaiTax(availabilityQuoteTax, tempAvailabilityQuoteTax);
                                tempAvailabilityQuoteTax = null;
                            }
                            if (_infant > 0)
                            {
                                tempAvailabilityQuoteTax = _flightRepository.GetFlightAvailabilityQuoteTax(_originRcd,
                                                                                                           _destinationRcd,
                                                                                                           _currencyRcd,
                                                                                                           "INF",
                                                                                                           _agencyCode,
                                                                                                           dtCurrent,
                                                                                                           _bookingDate);

                                FillAvaiTax(availabilityQuoteTax, tempAvailabilityQuoteTax);
                                tempAvailabilityQuoteTax = null;
                            }
                            if (string.IsNullOrEmpty(_otherPaxType) == false)
                            {
                                tempAvailabilityQuoteTax = _flightRepository.GetFlightAvailabilityQuoteTax(_originRcd,
                                                                                                           _destinationRcd,
                                                                                                           _currencyRcd,
                                                                                                           _otherPaxType,
                                                                                                           _agencyCode,
                                                                                                           dtCurrent,
                                                                                                           _bookingDate);

                                FillAvaiTax(availabilityQuoteTax, tempAvailabilityQuoteTax);
                                tempAvailabilityQuoteTax = null;
                            }

                            //increment day until reach _toDate
                            dtCurrent = dtCurrent.AddDays(1);
                        }

                        bool transitFlight = false;
                        foreach (Availability a in baseAvailability)
                        {
                            //Start fill tax to availability.
                            if (a.avl_show_net_total_flag == 1)
                            {
                                if (a.transit_flight_id.Equals(Guid.Empty))
                                {
                                    transitFlight = false;
                                }
                                else
                                {
                                    transitFlight = true;
                                }

                                if (_adult > 0)
                                {
                                    a.total_adult_fare = a.total_adult_fare + GetTaxAmount(availabilityQuoteTax,
                                                                                            a.departure_date,
                                                                                            a.total_adult_fare,
                                                                                            "ADULT",
                                                                                            a.fare_code,
                                                                                            a.currency_rcd,
                                                                                            a.booking_class_rcd,
                                                                                            _noVat,
                                                                                            transitFlight);
                                }
                                if (_child > 0)
                                {
                                    a.total_child_fare = a.total_child_fare + GetTaxAmount(availabilityQuoteTax,
                                                                                            a.departure_date,
                                                                                            a.total_adult_fare,
                                                                                            "CHD",
                                                                                            a.fare_code,
                                                                                            a.currency_rcd,
                                                                                            a.booking_class_rcd,
                                                                                            _noVat,
                                                                                            transitFlight);
                                }
                                if (_infant > 0)
                                {
                                    a.total_infant_fare = a.total_infant_fare + GetTaxAmount(availabilityQuoteTax,
                                                                                            a.departure_date,
                                                                                            a.total_adult_fare,
                                                                                            "INF",
                                                                                            a.fare_code,
                                                                                            a.currency_rcd,
                                                                                            a.booking_class_rcd,
                                                                                            _noVat,
                                                                                            transitFlight);
                                }
                                if (string.IsNullOrEmpty(_otherPaxType) == false)
                                {
                                    a.total_other_fare = a.total_other_fare + GetTaxAmount(availabilityQuoteTax,
                                                                                            a.departure_date,
                                                                                            a.total_adult_fare,
                                                                                            _otherPaxType,
                                                                                            a.fare_code,
                                                                                            a.currency_rcd,
                                                                                            a.booking_class_rcd,
                                                                                            _noVat,
                                                                                            transitFlight);
                                }
                                
                            }
                        }
                    }
                }

                return baseAvailability;
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        private void FillAvaiTax(IList<AvailabilityQuoteTax> container, 
                                IEnumerable<AvailabilityQuoteTax> value)
        {
            if (container != null && value != null)
            {
                foreach (AvailabilityQuoteTax at in value)
                {
                    container.Add(at);
                }
            }
        }

        private decimal GetTaxAmount(IList<AvailabilityQuoteTax> availabilityQuoteTax,
                                    DateTime departureDate,
                                    decimal fareAmount,
                                    string paxType,
                                    string fareBasis,
                                    string currencyRcd,
                                    string bookingClass,
                                    bool bNoVat,
                                    bool transitFlight)
        {
            decimal dTaxAmount = 0;
            decimal dTotalTaxAmount = 0;
            decimal transitAmount = 0;

            for (int i = 0; i < availabilityQuoteTax.Count; i++)
            {
                dTaxAmount = 0;
                if (availabilityQuoteTax[i].departure_date.Date == departureDate &&
                    availabilityQuoteTax[i].passenger_type_rcd == paxType)
                {
                    if (string.IsNullOrEmpty(availabilityQuoteTax[i].fare_code) == false &&
                        availabilityQuoteTax[i].fare_code == fareBasis)
                    {
                        //Have fare basis but not match.
                    }
                    else if (string.IsNullOrEmpty(availabilityQuoteTax[i].valid_for_class) == false &&
                            availabilityQuoteTax[i].valid_for_class.IndexOf(bookingClass) == -1)
                    {
                        //Have class mapping but fare class does not match.
                    }
                    else
                    {
                        if (bNoVat == true)
                        {
                            dTaxAmount = availabilityQuoteTax[i].tax_amount;
                        }
                        else
                        {
                            dTaxAmount = availabilityQuoteTax[i].tax_amount_incl;
                        }

                        //Do tax conversion if currency not match.
                        if (availabilityQuoteTax[i].tax_currency != currencyRcd)
                        {
                            if (availabilityQuoteTax[i].exchange_to_accounting > 0)
                            {
                                dTaxAmount = availabilityQuoteTax[i].exchange_to_accounting * dTaxAmount;
                            }

                            if (availabilityQuoteTax[i].exchange_from_accounting > 0)
                            {
                                dTaxAmount = availabilityQuoteTax[i].exchange_from_accounting * dTaxAmount;
                            }
                            dTaxAmount = Math.Round(dTaxAmount);
                        }

                        //Find percentage handling.
                        decimal percentageTax = 0;
                        if (availabilityQuoteTax[i].tax_percentage != 0)
                        {
                            if (fareAmount != 0)
                            {
                                percentageTax = fareAmount * (availabilityQuoteTax[i].tax_percentage / 100);
                                percentageTax = percentageTax + (percentageTax * (availabilityQuoteTax[i].vat_percentage / 100));
                            }

                            if (availabilityQuoteTax[i].minimum_tax_amount_flag == 0)
                            {
                                if (dTaxAmount < percentageTax)
                                {
                                    dTaxAmount = percentageTax;
                                }
                            }
                            else
                            {
                                if (dTaxAmount > percentageTax)
                                {
                                    dTaxAmount = percentageTax;
                                }
                            }
                        }

                        if (transitFlight == true)
                        {
                            //Transit information.
                            decimal otaxDistribution = Math.Round(dTaxAmount * availabilityQuoteTax[i].origin_tax_distribution / 100, 2);
                            decimal dtaxDistribution = Math.Round(dTaxAmount * availabilityQuoteTax[i].destination_tax_distribution / 100, 2);

                            transitAmount = 0;
                            transitAmount = transitAmount + otaxDistribution + dtaxDistribution;


                            if (availabilityQuoteTax[i].transit_tax_only_flag == 0)
                            {
                                dTotalTaxAmount = dTotalTaxAmount + transitAmount;
                            }
                        }
                        else
                        {
                            dTotalTaxAmount = dTotalTaxAmount + dTaxAmount;
                        }
                        
                    }
                }
            }

            return dTotalTaxAmount;
        }
    }
}
