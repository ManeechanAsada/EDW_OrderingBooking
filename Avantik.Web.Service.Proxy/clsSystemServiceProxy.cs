using Avantik.Web.Service.Contracts;
using Avantik.Web.Service.Message;
using Avantik.Web.Service.Message.System;
using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Text;

namespace Avantik.Web.Service.Proxy
{
    public class SystemServiceProxy : ClientBase<ISystemService>, ISystemService
    {
        //public OriginsResponse GetOrigins(OriginsRequest request)
        //{
        //    return base.Channel.GetOrigins(request);
        //}

        //public DestinationsResponse GetDestinations(DestinationsRequest request)
        //{
        //    return base.Channel.GetDestinations(request);
        //}

        //public GetCountryResponse GetCountry(GetCountryRequest request)
        //{
        //    return base.Channel.GetCountry(request);
        //}

        //public GetLanguageResponse GetLanguage(GetLanguageRequest request)
        //{
        //    return base.Channel.GetLanguage(request);
        //}
        //public GetTitleResponse GetTitle(GetTitleRequest request)
        //{
        //    return base.Channel.GetTitle(request);
        //}
        //public GetCurrencyResponse GetCurrency(GetCurrencyRequest request)
        //{
        //    return base.Channel.GetCurrency(request);
        //}
        //public GetSpecialServiceResponse GetSpecialService(GetSpecialServiceRequest request)
        //{
        //    return base.Channel.GetSpecialService(request);
        //}
        //public GetFormOfPaymentSubTypesResponse GetFormOfPaymentSubTypes(GetFormOfPaymentSubTypesRequest request)
        //{
        //    return base.Channel.GetFormOfPaymentSubTypes(request);
        //}

        //public DocumentResponse GetDocumentType(DocumentRequest request)
        //{
        //    return base.Channel.GetDocumentType(request);
        //}

    }
}
