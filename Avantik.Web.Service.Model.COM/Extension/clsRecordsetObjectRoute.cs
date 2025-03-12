using Avantik.Web.Service.COMHelper;
using Avantik.Web.Service.Entity.Route;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Avantik.Web.Service.Model.COM
{
    public static class RecordsetObjectRoute
    {
        public static void FillOrigin(this IList<Route> routes,ref ADODB.Recordset rs)
        {
            if (rs != null && rs.RecordCount > 0)
            {
                
                Route route = null;

                try
                {
                    rs.MoveFirst();
                    while (!rs.EOF)
                    {
                        route = new Route();

                        route.OriginRcd = RecordsetHelper.ToString(rs, "origin_rcd");
                        route.DisplayName = RecordsetHelper.ToString(rs, "display_name");
                        route.CountryRcd = RecordsetHelper.ToString(rs, "country_rcd");
                        route.CurrencyRcd = RecordsetHelper.ToString(rs, "currency_rcd");
                        route.RoutesTot = RecordsetHelper.ToByte(rs, "routes_tot");
                        route.RoutesAvl = RecordsetHelper.ToByte(rs, "routes_avl");
                        route.RoutesB2C = RecordsetHelper.ToByte(rs, "routes_b2c");
                        route.RoutesB2B = RecordsetHelper.ToByte(rs, "routes_b2b");
                        route.RoutesB2S = RecordsetHelper.ToByte(rs, "routes_b2s");
                        route.RoutesAPI = RecordsetHelper.ToByte(rs, "routes_api");
                        route.RoutesB2T = RecordsetHelper.ToByte(rs, "routes_b2t");
                        routes.Add(route);
                        rs.MoveNext();
                    }

                }
                catch
                {
                    throw;
                }
            }
          
        }

        public static void FillDestination(this  IList<Route> routes, ref ADODB.Recordset rs)
        {
            if (rs != null && rs.RecordCount > 0)
            {
                Route route = null;

                try
                {
                    rs.MoveFirst();
                    while (!rs.EOF)
                    {
                        route = new Route();

                        route.OriginRcd = RecordsetHelper.ToString(rs, "origin_rcd");
                        route.DestinationRcd = RecordsetHelper.ToString(rs, "destination_rcd");
                        route.SegmentChangeFeeFlag = RecordsetHelper.ToBoolean(rs, "segment_change_fee_flag");
                        route.TransitFlag = RecordsetHelper.ToBoolean(rs, "transit_flag");
                        route.DirectFlag = RecordsetHelper.ToBoolean(rs, "direct_flag");
                        route.AvlFlag = RecordsetHelper.ToBoolean(rs, "avl_flag");
                        route.B2CFlag = RecordsetHelper.ToBoolean(rs, "b2c_flag");
                        route.B2BFlag = RecordsetHelper.ToBoolean(rs, "b2b_flag");
                        route.B2TFlag = RecordsetHelper.ToBoolean(rs, "b2t_flag");
                        route.DayRange = RecordsetHelper.ToInt16(rs, "day_range");
                        route.ShowRedressNumberFlag = RecordsetHelper.ToBoolean(rs, "show_redress_number_flag");
                        route.RequirePassengerTitleFlag = RecordsetHelper.ToBoolean(rs, "require_passenger_title_flag");
                        route.RequirePassengerGenderFlag = RecordsetHelper.ToBoolean(rs, "require_passenger_gender_flag");
                        route.RequireDateOfBirthFlag = RecordsetHelper.ToBoolean(rs, "require_date_of_birth_flag");
                        route.RequireDocumentDetailsFlag = RecordsetHelper.ToBoolean(rs, "require_document_details_flag");
                        route.RequirePassengerWeightFlag = RecordsetHelper.ToBoolean(rs, "require_passenger_weight_flag");
                        route.SpecialServiceFeeFlag = RecordsetHelper.ToBoolean(rs, "special_service_fee_flag");
                        route.ShowInsuranceOnWebFlag = RecordsetHelper.ToBoolean(rs, "show_insurance_on_web_flag");
                        route.DisplayName = RecordsetHelper.ToString(rs, "display_name");
                        route.CountryRcd = RecordsetHelper.ToString(rs, "destination_country_rcd");
                        route.CurrencyRcd = RecordsetHelper.ToString(rs, "currency_rcd");

                        routes.Add(route);
                        rs.MoveNext();
                    }

                }
                catch
                {
                    throw;
                }
            }
        }
    }
}
