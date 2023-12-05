﻿using RpaLib.APIs;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using RpaLib.APIs.GraphQL;
using RpaLib.APIs.Pipefy.Model;
using RpaLib.APIs.Pipefy.Exception;
using RpaLib.Tracing;
using RpaLib.Exceptions;
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
            Uri = uri ?? PipefyDefaultEndPoint;
            
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

        #region Queries

        public GraphQlResponse<MeResult> QueryUserInfo()
        {
            string query = @"
            {
                me {
                    id
                    email
                }
            }";

            return GraphQl.Query<MeResult>(query, Uri, Token, JsonSerializerSettings);
        }

        public GraphQlResponse<PipeResult> QueryPhases() => QueryPhases(PipeId);
        public GraphQlResponse<PipeResult> QueryPhases(string pipeId)
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

            return GraphQl.Query<PipeResult>(query, Uri, Token, JsonSerializerSettings);
        }

        public GraphQlResponse<PhaseResult> QueryPhaseCards(string phaseId, int max = 30, string afterCursor = null)
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
                      current_phase {
                        id
                      }
                      createdAt
                      age
            	      fields {
            	        name
            		    phase_field { id }
                        field {
                          id
                          options
                          label
                          type
                        }
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

            return GraphQl.Query<PhaseResult>(query, Uri, Token, JsonSerializerSettings);
        }

        public enum OrderCardsBy
        {
            Nothing,
            Newer,
            Older
        }

        public GraphQlResponse<PhaseResult> QueryAllPhaseCards(string phaseId, int limit = 0, OrderCardsBy orderBy = OrderCardsBy.Nothing)
        {
            List<CardEdge> cardEdges = new List<CardEdge>();
            GraphQlResponse<PhaseResult> response;

            string pageCursor = null;
            bool hasNextPage = false;
            bool limitReached = false;

            do
            {
                response = QueryPhaseCards(phaseId, afterCursor: pageCursor);

                var phaseCards = response.Data.Phase.Cards.Edges;
                hasNextPage = response.Data.Phase.Cards.Pageinfo.Hasnextpage;

                if (hasNextPage)
                    pageCursor = response.Data.Phase.Cards.Pageinfo.Endcursor;

                if (phaseCards.Count == 0)
                {
                    Trace.WriteLine($"No cards in phase \"[ID: {phaseId}]\". Total cards: {phaseCards.Count}");
                    return null;
                }

                foreach (var cardEdge in phaseCards)
                {
                    cardEdges.Add(cardEdge);

                    // limit results
                    if (limit > 0)
                        if (cardEdges.Count == limit)
                        {
                            limitReached = true;
                            break;
                        }
                
                }

            } while (hasNextPage && !limitReached);

            // order results
            if (orderBy == OrderCardsBy.Older)
            {
                cardEdges = cardEdges.OrderBy(x => x.Node.Createdat).ToList();
            }
            else if (orderBy == OrderCardsBy.Newer)
            {
                cardEdges = cardEdges.OrderBy(x => x.Node.Createdat).Reverse().ToList();
            }

            response.Data.Phase.Cards.Edges = cardEdges;

            return response;
        }

        public GraphQlResponse<PhaseResult> QueryPhaseFields(string phaseId)
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

            return GraphQl.Query<PhaseResult>(query, Uri, Token, JsonSerializerSettings);
        }

        public GraphQlResponse<CardResult> QueryCard(string cardId)
        {
            string query = @"
            {
              card (id: $cardId) {
                id
                current_phase {
                  id
                }
                age
                createdAt
                expiration {
                  expiredAt
                  shouldExpireAt
                }
                fields {
                  name
                  phase_field { id }
                  field {
                    id
                    options
                    label
                    type
                  }
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

            return GraphQl.Query<CardResult>(query, Uri, Token, JsonSerializerSettings);
        }

        #endregion

        #region Mutations
        /// <summary>
        /// Move a card to a phase.
        /// </summary>
        /// <param name="cardId">The ID of a card. Can be string or int literal in GraphQL.</param>
        /// <param name="destPhaseId">The ID of the destination phase. Can be string or int literal in GraphQL.</param>
        /// <returns>A GraphQlResponse with Data of type CardQuery containing the new card phase.</returns>
        public GraphQlResponse<CardResult> MoveCardToPhase(string cardId, string destPhaseId)
        {
            var query = @"
                mutation {
                  moveCardToPhase(input: {
                    card_id: ""<<IdCard>>"",
                    destination_phase_id: ""<<IdPhase>>""
                  }) {
                    card {
                      current_phase {
                        id
                      }
                    }
                  }
                }".Replace("<<IdCard>>", cardId).Replace("<<IdPhase>>", destPhaseId);

            return GraphQl.Query<CardResult>(query, Uri, Token);
        }

        public GraphQlResponse<UpdateCardFieldResult> UpdateCardField(string cardId, string fieldId, string newValue)
        {
            var escapedNewValue = newValue.Replace(@"\", @"\\").Replace(@"""", @"\""");
            var query = @"
                mutation {
                    updateCardField(input: { 
                        card_id: ""<<IdCard>>"",
                        field_id: ""<<IdField>>"",
                        new_value: ""<<NewValue>>""
                    }) {
                        success
                        card {
                            id
                            fields {
                                name
                                value
                                updated_at
                            }
                        }
                    }
                }".Replace("<<IdCard>>", cardId).Replace("<<IdField>>", fieldId).Replace("<<NewValue>>", escapedNewValue);

            return GraphQl.Query<UpdateCardFieldResult>(query, Uri, Token);
        }

        #endregion

        /// <summary>
        /// Query pipefy info using API calls and redirects the info output to the Tracing.
        /// </summary>
        /// <param name="infoType">A PipefyInfo enum type indicating the kind of information to retrieve.</param>
        /// <param name="phaseId">The Pipefy phase ID that is needed to perform some queries.</param>
        /// <exception cref="ArgumentNullException">Thrown when a parameter is needed by some query but wasn't provided.</exception>
        public void ShowInfo(PipefyInfo infoType, string phaseId = null, string cardId = null)
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
                        throw new RpaLibArgumentNullException("Phase ID argument is mandatory to fetch the fields from that phase.");
                    PhaseResult phaseQuery = QueryPhaseFields(phaseId).Data;
                    Trace.WriteLine($"Phase (ID {phaseQuery.Phase.Id}) {phaseQuery.Phase.Name}, has the following fields:");
                    foreach (PhaseField field in phaseQuery.Phase.Fields)
                    {
                        string msg = string.Join("\n",
                            "    ID: " + field.Id,
                            "    Label: " + field.Label,
                            "    Options: " + string.Join(", ", field.Options));
                        Trace.WriteLine(msg, withTimeSpec: false);
                    }
                    break;

                case PipefyInfo.CardFields:
                    if (cardId == null)
                        throw new RpaLibArgumentNullException("Card ID argument is mandatory to fetch the fields from that card.");
                    CardResult cardQuery = QueryCard(cardId).Data;
                    Trace.WriteLine($"Card (ID {cardQuery.Card.Id}) has the following fields");
                    foreach (CardField field in cardQuery.Card.Fields)
                    {
                        string msg = string.Join(Environment.NewLine,
                            "    ID: " + field.Field.Id,
                            "    Name: " + field.Name,
                            "    Options: " + string.Join(", ", field.Field.Options));
                        Trace.WriteLine(msg, withTimeSpec: false);
                    }
                    break;
            }
        }

        /// <summary>
        /// Search for a card field which name is equal to "fieldName" parameter. If more than one field is found with the same name
        /// the first one in the sequence of found fields will be returned.
        /// </summary>
        /// <param name="card">The card object to look for.</param>
        /// <param name="fieldName">The card field name to search.</param>
        /// <returns></returns>
        /// <exception cref="CardFieldNotFoundException">Thrown when the card could not be found by the name passed as argument.</exception>
        public static string GetFieldValue(Card card, string fieldName)
        {
            string foundFieldValue = null;
            var foundFields = from field in card.Fields
                              where field.Name == fieldName
                              select field.Value;

            if (foundFields.Count() <= 0)
            {
                throw new CardFieldNotFoundException(card.Id, fieldName, foundFields.Count());
            }
            else if (foundFields.Count() > 0)
            {
                Trace.WriteLine($"Number of fields found: {foundFields.Count()}. The first (index 0) will be selected.");
                foundFieldValue = foundFields.FirstOrDefault();
            }

            return foundFieldValue;
        }
    }
}
