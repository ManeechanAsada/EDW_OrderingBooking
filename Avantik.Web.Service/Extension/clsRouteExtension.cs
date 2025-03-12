using Avantik.Web.Service.Entity.Route;
using Avantik.Web.Service.Message;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Avantik.Web.Service.Extension
{
    public static class RouteExtension
    {
        public static IEnumerable<RouteView> MappingOriginsView(this IEnumerable<Route> routes)
        {
            try
            {
                if (routes != null)
                {
                    IList<RouteView> rts = new List<RouteView>();
                    RouteView rv = null;
                    foreach (Route a in routes)
                    {
                        rv = new RouteView();


                        rv.origin_rcd = a.OriginRcd;
                        rv.display_name = a.DisplayName;
                        rv.country_rcd = a.CountryRcd;
                        rv.currency_rcd = a.CurrencyRcd;
                        rv.routes_tot = a.RoutesTot;
                        rv.routes_avl = a.RoutesAvl;
                        rv.routes_b2c = a.RoutesB2C;
                        rv.routes_b2b = a.RoutesB2B;
                        rv.routes_b2s = a.RoutesB2S;
                        rv.routes_api = a.RoutesAPI;
                        rv.routes_b2t = a.RoutesB2T;

                        rts.Add(rv);

                    }

                    return rts;
                }
                else
                {
                    return null;
                }
            }
            catch
            {
                throw;
            }
        }

        public static IList<RouteView> MappingDestinationsView(this IList<Route> routes)
        {
            try
            {
                if (routes != null)
                {
                    IList<RouteView> rts = new List<RouteView>();
                    RouteView rv = null;
                    foreach (Route a in routes)
                    {
                        rv = new RouteView();

                        rv.origin_rcd = a.OriginRcd;
                        rv.destination_rcd = a.DestinationRcd;
                        rv.segment_change_fee_flag = a.SegmentChangeFeeFlag;
                        rv.transit_flag = a.TransitFlag;
                        rv.direct_flag = a.DirectFlag;
                        rv.avl_flag = a.AvlFlag;
                        rv.b2c_flag = a.B2CFlag;
                        rv.b2b_flag = a.B2BFlag;
                        rv.b2t_flag = a.B2TFlag;
                        rv.day_range = a.DayRange;
                        rv.show_redress_number_flag = a.ShowRedressNumberFlag;
                        rv.require_passenger_title_flag = a.RequirePassengerTitleFlag;
                        rv.require_passenger_gender_flag = a.RequirePassengerGenderFlag;
                        rv.require_date_of_birth_flag = a.RequireDateOfBirthFlag;
                        rv.require_document_details_flag = a.RequireDocumentDetailsFlag;
                        rv.require_passenger_weight_flag = a.RequirePassengerWeightFlag;
                        rv.show_special_service_on_web_flag = a.SpecialServiceFeeFlag;
                        rv.show_insurance_on_web_flag = a.ShowInsuranceOnWebFlag;
                        rv.display_name = a.DisplayName;
                        rv.country_rcd = a.CountryRcd;
                        rv.currency_rcd = a.CurrencyRcd;

                        rts.Add(rv);

                    }

                    return rts;
                }
                else
                {
                    return null;
                }
            }
            catch
            {
                throw;
            }
        }
    }
}
