using RpaLib.APIs;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using RpaLib.APIs.GraphQL;
using RpaLib.APIs.Pipefy.Model;
using RpaLib.Tracing;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;

namespace RpaLib.APIs.Pipefy
{
    public class Pipefy
    {
        public readonly static string PipefyDefaultEndPoint = @"https://app.pipefy.com/queries";
        public string Token { get; private set; }
        public string PipeId { get; private set; }
        public string Uri { get; private set; }
        public JsonSerializerSettings JsonSerializerSettings { get; private set; }

        public Pipefy (string apiToken, string pipeId, string uri = null)
        {
            if (string.IsNullOrEmpty(uri))
            {
                Uri = PipefyDefaultEndPoint;
            }
            
            PipeId = pipeId;
            Token = apiToken;

            JsonSerializerSettings = CreateSnakeCaseJsonSerializerSettings();
        }

        public static JsonSerializerSettings CreateSnakeCaseJsonSerializerSettings()
        {
            var contractResolver = new DefaultContractResolver
            {
                NamingStrategy = new SnakeCaseNamingStrategy()
            };

            var jsonSerializerSettings = new JsonSerializerSettings
            {
                ContractResolver = contractResolver,
                Formatting = Formatting.Indented,
            };

            return jsonSerializerSettings;
        }

        public GraphQlResponse<MeQuery> QueryUserInfo()
        {
            string query = @"
            {
                me {
                    id
                    email
                }
            }";

            return GraphQl.Query<MeQuery>(query, Uri, Token, JsonSerializerSettings);
        }

        public GraphQlResponse<PipeQuery> QueryPhases() => QueryPhases(PipeId);
        public GraphQlResponse<PipeQuery> QueryPhases(string pipeId)
        {
            string query = @"
            {
                pipe (id: $pipeId) {
                    phases {
                        name
                        id
                        cards_count
                    }
                }
            }".Replace("$pipeId", pipeId);

            return GraphQl.Query<PipeQuery>(query, Uri, Token, JsonSerializerSettings);
        }

        public GraphQlResponse<PhaseQuery> QueryPhaseCards(string phaseId, int max = 30, string afterCursor = null)
        {
            string query = @"
            {
              phase (id: $phaseId) {
                cards_count
                cards (first: $max, after: ""$afterCursor"") {
                  pageInfo {
                    hasPreviousPage
            	    hasNextPage
            	    startCursor
            	    endCursor
                  }
                  edges {
                    cursor
                    node {
            	      id
            	      fields {
            	        name
            		phase_field { id }
            	        value
                      }
            	      labels {
            	        id
            	        name
            	        color
            	      }
            	    }
                  }
                }
              }
            }
            ".Replace("$phaseId", phaseId).Replace("$max", max.ToString());

            if (string.IsNullOrEmpty(afterCursor))
            {
                query = query.Replace(@", after: ""$afterCursor""", string.Empty);
            }
            else
            {
                query = query.Replace("$afterCursor", afterCursor);
            }

            return GraphQl.Query<PhaseQuery>(query, Uri, Token, JsonSerializerSettings);
        }

        public GraphQlResponse<PhaseQuery> QueryPhaseFields(string phaseId)
        {
            string query = @"
            {
              phase (id: $phaseId) {
                id
                name
                fields {
                  label
                  id
                  options
                }
              }
            }
            ".Replace("$phaseId", phaseId);

            return GraphQl.Query<PhaseQuery>(query, Uri, Token, JsonSerializerSettings);
        }

        public GraphQlResponse<Card> QueryCard(string cardId)
        {
            string query = @"
            {
              card (id: $cardId) {
                id
                current_phase {
                  id
                }
                age
                expiration {
                  expiredAt
                  shouldExpireAt
                }
                fields {
                  name
                  phase_field { id }
                  value
                }
                labels {
                  id
                  name
                  color
                }
                done
                due_date
                expired
                updated_at
              }
            }
            ".Replace("$cardId", cardId);

            return GraphQl.Query<Card>(query, Uri, Token, JsonSerializerSettings);
        }

        public void ShowInfo(PipefyInfo infoType, string phaseId = null)
        {
            switch (infoType)
            {
                case PipefyInfo.PhasesAndCardsCount:
                    var phases = QueryPhases().Data.Pipe.Phases;

                    foreach (Phase phase in phases)
                    {
                        Trace.WriteLine($"Number of cards in phase (ID {phase.Id}) \"{phase.Name}\": {phase.CardsCount}");
                    }
                    break;

                case PipefyInfo.PhaseFields:
                    if (phaseId == null)
                        throw new ArgumentNullException("Phase ID argument is mandatory to fetch the fields from that phase.");
                    PhaseQuery phaseQuery = QueryPhaseFields(phaseId).Data;
                    Trace.WriteLine($"Phase (ID {phaseQuery.Phase.Id}) {phaseQuery.Phase.Name}, has the following fields:");
                    foreach (PhaseField field in phaseQuery.Phase.Fields)
                    {
                        string msg = string.Join("\n",
                            string.Empty,
                            "    ID: " + field.Id,
                            "    Label: " + field.Label,
                            "    Options: " + string.Join(", ", field.Options));
                        Trace.WriteLine(msg);
                    }
                    break;
            }
        }
    }
}
