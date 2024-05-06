using GraphQL;
using GraphQL.Client.Http;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.Net.Http.Headers;

namespace RpaLib.APIs.GraphQL
{
    public class GraphQlResponse<T>
    {
        public string Queried { get; private set; }
        public T Data { get; private set; }
        public GraphQLError[] Errors { get; private set; }
        public Map Extensions { get; private set; }
        public HttpResponseHeaders ResponseHeaders { get; private set; }
        public HttpStatusCode StatusCode { get; private set; }
        public string Json { get => ToJson(); }
        public bool IsSuccessStatusCode
        {
            get => StatusCode == HttpStatusCode.OK;
        }

        public GraphQlResponse(GraphQLResponse<T> graphQLResponseObj, string queried)
        {
            Data = graphQLResponseObj.Data;
            Errors = graphQLResponseObj.Errors;
            Extensions = graphQLResponseObj.Extensions;
            ResponseHeaders = graphQLResponseObj.AsGraphQLHttpResponse().ResponseHeaders;
            StatusCode = graphQLResponseObj.AsGraphQLHttpResponse().StatusCode;
            Queried = queried;
        }

        public string ToJson()
        {
            return JsonConvert.SerializeObject(Data, Formatting.Indented);
        }

        public override string ToString() => ToJson();
    }
}
