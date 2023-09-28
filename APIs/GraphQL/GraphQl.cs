using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;
using GraphQL;
using GraphQL.Client.Http;
using GraphQL.Client.Serializer.Newtonsoft;
using Newtonsoft.Json;

namespace RpaLib.APIs.GraphQL
{
    public static class GraphQl
    {
        private static GraphQLHttpClient Connect(string endpoint, string accessToken, JsonSerializerSettings jsonSerializerSettings = null)
        {
            var newtonsoftJsonSerializer = jsonSerializerSettings == null ? new NewtonsoftJsonSerializer() : new NewtonsoftJsonSerializer(jsonSerializerSettings);

            var graphQLHttpClientOptions = new GraphQLHttpClientOptions
            {
                EndPoint = new Uri(endpoint)
            };

            var httpClient = new HttpClient();

            httpClient.DefaultRequestHeaders.Add("Authorization", accessToken);

            var client = new GraphQLHttpClient(
                graphQLHttpClientOptions,
                newtonsoftJsonSerializer,
                httpClient);

            return client;
        }

        public static GraphQlResponse<dynamic> Query(string graphQlQuery, string endpoint, string accessToken, JsonSerializerSettings jsonSerializerSettings = null)
        {
            return Query<dynamic>(graphQlQuery, endpoint, accessToken);
        }

        public static GraphQlResponse<T> Query<T>(string graphQlQuery, string endpoint, string accessToken, JsonSerializerSettings jsonSerializerSettings = null)
        {
            return Query<T>(graphQlQuery, Connect(endpoint, accessToken, jsonSerializerSettings));
        }

        public static GraphQlResponse<T> Query<T>(string graphQlQuery, GraphQLHttpClient graphQLHttpClient)
        {
            var queryAsyncTask = QueryAsync<T>(graphQlQuery, graphQLHttpClient);

            Task.WaitAll(queryAsyncTask);

            var result = queryAsyncTask.Result;

            return result;
        }

        public static async Task<GraphQlResponse<dynamic>> QueryAsync(string grpahQlQuery, string endpoint, string accessToken, JsonSerializerSettings jsonSerializerSettings = null)
        {
            return await QueryAsync<dynamic>(grpahQlQuery, Connect(endpoint, accessToken, jsonSerializerSettings));
        }

        public static async Task<GraphQlResponse<dynamic>> QueryAsync(string grpahQlQuery, GraphQLHttpClient graphQLHttpClient)
        {
            return await QueryAsync<dynamic>(grpahQlQuery, graphQLHttpClient);
        }

        public static async Task<GraphQlResponse<T>> QueryAsync<T>(string graphQlQuery, string endpoint, string accessToken, JsonSerializerSettings jsonSerializerSettings = null)
        {
            return await QueryAsync<T>(graphQlQuery, Connect(endpoint, accessToken, jsonSerializerSettings)); ;
        }

        public static async Task<GraphQlResponse<T>> QueryAsync<T>(string graphQlQuery, GraphQLHttpClient graphQLHttpClient)
        {
            var request = new GraphQLRequest
            {
                Query = graphQlQuery,
            };

            var response = await graphQLHttpClient.SendQueryAsync<T>(request);

            graphQLHttpClient.Dispose();

            return new GraphQlResponse<T>(response);
        }

    }
}
