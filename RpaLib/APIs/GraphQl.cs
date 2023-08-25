using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using GraphQL;
using GraphQL.Client.Http;
using GraphQL.Client.Serializer.Newtonsoft;

namespace RpaLib.APIs
{
    public class GraphQl
    {
        public string AccessToken { get; private set; }
        public string Endpoint { get; private set; }

        private GraphQLHttpClient GraphQLClient { get; set; }

        public GraphQl(string endpoint, string accessToken)
        {
            AccessToken = accessToken;
            Endpoint = endpoint;

            var graphQLHttpClientOptions = new GraphQLHttpClientOptions
            {
                EndPoint = new Uri(endpoint)
            };

            var httpClient = new HttpClient();

            httpClient.DefaultRequestHeaders.Add("Authorization", accessToken);

            GraphQLClient = new GraphQLHttpClient(
                graphQLHttpClientOptions,
                new NewtonsoftJsonSerializer(),
                httpClient);
        }

        public GraphQlResponse Query (string graphQlQuery)
        {
            var queryAsyncTask = QueryAsync(graphQlQuery);

            Task.WaitAll(queryAsyncTask);

            var result = queryAsyncTask.Result;

            return result;
        }

        public async Task<GraphQlResponse> QueryAsync (string graphQlQuery)
        {
            var request = new GraphQLRequest
            {
                Query = graphQlQuery,
            };

            var response = await GraphQLClient.SendQueryAsync<dynamic>(request);

            return new GraphQlResponse(response);
        }
    }
}
