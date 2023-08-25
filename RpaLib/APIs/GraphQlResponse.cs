using GraphQL;
using GraphQL.Client.Http;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RpaLib.APIs
{
    public class GraphQlResponse
    {
        public dynamic Data { get; private set; }
        public dynamic Errors { get; private set; }
        public dynamic Extensions { get; private set; }
        public dynamic ResponseHeaders { get; private set; }
        public dynamic StatusCode { get; private set; }

        public GraphQlResponse(GraphQLResponse<dynamic> graphQLResponseObj)
        {
            Data = graphQLResponseObj.Data;
            Errors = graphQLResponseObj.Errors;
            Extensions = graphQLResponseObj.Extensions;
            ResponseHeaders = graphQLResponseObj.AsGraphQLHttpResponse().ResponseHeaders;
            StatusCode = graphQLResponseObj.AsGraphQLHttpResponse().StatusCode;
        }

        public string ToJson()
        {
            return JsonConvert.SerializeObject(Data, Formatting.Indented);
        }
    }
}
