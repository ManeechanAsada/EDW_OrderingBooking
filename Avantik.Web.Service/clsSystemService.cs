using Avantik.Web.Service.Contracts;
using Avantik.Web.Service.Entity.Route;
using Avantik.Web.Service.Message;
using Avantik.Web.Service.Model.Contract;
using Avantik.Web.Service.Model.Factory;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Avantik.Web.Service.Extension;
using Avantik.Web.Service.Helpers;
using System.IO;
using System.Runtime.Serialization.Json;
using Avantik.Web.Service.Model;
using Avantik.Web.Service.Extension.System;
using Avantik.Web.Service.Message.System;
using Avantik.Web.Service.Entity;
using System.ServiceModel;

namespace Avantik.Web.Service
{
    [ServiceBehavior(AddressFilterMode = AddressFilterMode.Any)]
    public class SystemService:ISystemService
    {
        public OriginsResponse GetOrigins(OriginsRequest Request)
        {
            IRouteService objRouteService = RouteServiceFactory.CreateInstance();
            OriginsResponse respose = new OriginsResponse();

            try
            {
                IEnumerable<Route> OriginsResult = objRouteService.GetOrigins(Request.Language,Request.B2CFlag, Request.B2BFlag,Request.B2EFlag,Request.B2SFlag,Request.APIFlag);

                //string xxx = XMLHelper.JsonSerializer(typeof(OriginsRequest), Request);

                if (OriginsResult != null)
                {
                    respose.Success = true;
                    respose.Message = "Success request.";
                    respose.Routes = OriginsResult.MappingOriginsView();
                }
                else
                {
                    respose.Success = false;
                    respose.Message = "No Route result return.";
                }
            }
            catch (System.Exception ex)
            {
                respose.Success = false;
                respose.Message = ex.Message;

                //Logger.Instance(Logger.LogType.Mail).WriteLog(ex, XMLHelper.JsonSerializer(typeof(OriginsRequest), Request));
            }
            return respose;
        }

        public DestinationsResponse GetDestinations(DestinationsRequest Request)
        {
            IRouteService objRouteService = RouteServiceFactory.CreateInstance();
            DestinationsResponse respose = new DestinationsResponse();

            try
            {
                IList<Route> DestinationsResult = objRouteService.GetDestinations(Request.Language, Request.B2CFlag, Request.B2BFlag, Request.B2EFlag, Request.B2SFlag, Request.APIFlag);


                if (DestinationsResult != null)
                {
                    respose.Success = true;
                    respose.Message = "Success request.";
                    respose.Routes = DestinationsResult.MappingDestinationsView();
                }
                else
                {
                    respose.Success = false;
                    respose.Message = "No Route result return.";
                }

            }
            catch (System.Exception ex)
            {
                respose.Success = false;
                respose.Message = ex.Message;

                //Logger.Instance(Logger.LogType.Mail).WriteLog(ex, XMLHelper.JsonSerializer(typeof(DestinationsRequest), Request));
            }
            return respose;
        }

        public GetCountryResponse GetCountry(GetCountryRequest Request)
        {
            ISystemModelService objSystemService = SystemServiceFactory.CreateInstance();
            GetCountryResponse response = new GetCountryResponse();

            try
            {
                IList<Entity.Country> countries = objSystemService.GetCountry(Request.StrLanguage);

                if (countries != null)
                {
                    // map to response
                    response.Countries = countries.ToObjMessage();
                    response.Message = "Success";
                    response.Success = true;
                }
                else
                {
                    response.Message = "Fail";
                    response.Success = false;
                }
            }
            catch
            {
                response.Message = "Get Country error.";
                response.Success = false;
            }

            return response;
        }
        public GetLanguageResponse GetLanguage(GetLanguageRequest Request)
        {
            ISystemModelService objSystemService = SystemServiceFactory.CreateInstance();
            GetLanguageResponse response = new GetLanguageResponse();

            try
            {
                IList<Entity.Language> language = objSystemService.GetLanguage(Request.StrLanguage);

                if (language != null)
                {
                    // map to response
                    response.Languages = language.ToObjMessage();
                    response.Message = "Success";
                    response.Success = true;
                }
                else
                {
                    response.Message = "Fail";
                    response.Success = false;
                }
            }
            catch
            {
                response.Message = "Get Language error.";
                response.Success = false;
            }

            return response;
        }
        public GetTitleResponse GetTitle(GetTitleRequest Request)
        {
            ISystemModelService objSystemService = Model.SystemServiceFactory.CreateInstance();
            GetTitleResponse response = new GetTitleResponse();

            try
            {
                IList<Entity.Title> titles = objSystemService.GetTitle(Request.StrLanguage);

                if (titles != null)
                {
                    // map to response
                    response.Titles = titles.ToObjMessage();
                    response.Message = "Success";
                    response.Success = true;
                }
                else
                {
                    response.Message = "Fail";
                    response.Success = false;
                }
            }
            catch
            {
                response.Message = "Get title error.";
                response.Success = false;
            }

            return response;
        }
        public GetCurrencyResponse GetCurrency(GetCurrencyRequest Request)
        {
            ISystemModelService objSystemService = SystemServiceFactory.CreateInstance();
            GetCurrencyResponse response = new GetCurrencyResponse();

            try
            {
                IList<Entity.Currency> currencies = objSystemService.GetCurrency(Request.StrLanguage);

                if (currencies != null)
                {
                    // map to response
                    response.Currencies = currencies.ToObjMessage();
                    response.Message = "Success";
                    response.Success = true;
                }
                else
                {
                    response.Message = "Fail";
                    response.Success = false;
                }
            }
            catch
            {
                response.Message = "Get Currency error.";
                response.Success = false;
            }

            return response;
        }
        public GetSpecialServiceResponse GetSpecialService(GetSpecialServiceRequest Request)
        {
            ISystemModelService objSystemService = SystemServiceFactory.CreateInstance();
            GetSpecialServiceResponse response = new GetSpecialServiceResponse();

            try
            {
                IList<Entity.SpecialService> specialServices = objSystemService.GetSpecialService(Request.StrLanguage);

                if (specialServices != null)
                {
                    // map to response
                    response.SpecialServices = specialServices.ToObjMessage();
                    response.Message = "Success";
                    response.Success = true;
                }
                else
                {
                    response.Message = "Fail";
                    response.Success = false;
                }
            }
            catch
            {
                response.Message = "Get SpecialServices error.";
                response.Success = false;
            }

            return response;
        }
        public GetFormOfPaymentSubTypesResponse GetFormOfPaymentSubTypes(GetFormOfPaymentSubTypesRequest Request)
        {
            ISystemModelService objSystemService = SystemServiceFactory.CreateInstance();
            GetFormOfPaymentSubTypesResponse response = new GetFormOfPaymentSubTypesResponse();

            try
            {
                Entity.FormOfPayment.FormOfPayment formOfPayment = objSystemService.GetFormOfPaymentSubType(Request.Type, Request.StrLanguage);

                if (formOfPayment != null)
                {
                    // map to response
                    response.FormOfPayment = formOfPayment.ToObjMessage();
                    response.Message = "Success";
                    response.Success = true;
                }
                else
                {
                    response.Message = "Fail";
                    response.Success = false;
                }
            }
            catch
            {
                response.Message = "Get FormOfPaymentSubTypes error.";
                response.Success = false;
            }

            return response;
        }

        public DocumentResponse GetDocumentType(DocumentRequest Request)
        {
            ISystemModelService objDocumentService = SystemServiceFactory.CreateInstance();
            DocumentResponse respose = new DocumentResponse();

            try
            {
                IEnumerable<Document> DocumentTypeResult = objDocumentService.GetDocumentType(Request.Language);


                if (DocumentTypeResult != null)
                {
                    respose.Success = true;
                    respose.Message = "Success request.";
                    respose.DocumentType = DocumentTypeResult.MappingDocumentTypeView();
                }
                else
                {
                    respose.Success = false;
                    respose.Message = "No Document result return.";
                }

            }
            catch (System.Exception ex)
            {
                respose.Success = false;
                respose.Message = ex.Message;
               // Logger.Instance(Logger.LogType.Mail).WriteLog(ex, XMLHelper.JsonSerializer(typeof(DocumentRequest), Request));
            }
            return respose;
        }

    }
}
