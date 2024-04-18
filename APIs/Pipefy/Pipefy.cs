using RpaLib.APIs;
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
using System.IO;
using RpaLib.ProcessAutomation;

namespace RpaLib.APIs.Pipefy
{
    public class Pipefy
    {
        public readonly static string PipefyDefaultEndPoint = @"https://app.pipefy.com/queries";
        public string Token { get; private set; }
        public string PipeId { get; private set; }
        public string Uri { get; private set; }
        public JsonSerializerSettings JsonSerializerSettingsSnake { get; private set; }
        public JsonSerializerSettings JsonSerializerSettingsCamel { get; private set; }

        #region GraphQl_queries
        private string queryUserInfo = @"
            {
                me {
                    id
                    email
                    name
                    username
                }
            }";

        private string queryPipe = @"
            {
                pipe(id: ""<<PipeId>>"") {
                    organizationId
                    name
                    noun
                    labels {
                        id
                        name
                        color
                    }
                    phases {
                        id
                        name
                        cards {
                            edges {
                                cursor
                                node {
                                    id
                                    age
                                    createdAt
                                    createdBy {
                                        name
                                        email
                                        id
                                        phone
                                        username
                                    }
                                    done
                                    fields {
                                        name
                                        field {
                                            id
                                            label
                                            options
                                            type
                                            description
                                            help
                                            uuid
                                        }
                                    }
                                }
                            }
                            pageInfo {
                                endCursor
                                hasNextPage
                                hasPreviousPage
                                startCursor
                            }
                        }
                        fields {
                            id
                            label
                            options
                            type
                            required
                            editable
                            description
                        }
                    }
                }
            }";

        private string queryPipeSimple = @"
            {
                pipe(id: ""<<PipeId>>"") {
                    organizationId
                    name
                    noun
                    labels {
                        id
                        name
                        color
                    }
                }
            }
            ";

        private string queryOrganizations = @"
            {
                organizations {
                    createdAt
                    createdBy {
                        id
                        name
                        email
                    }
                    customLogoUrl
                    freemium
                    id
                    membersCount
                    name
                    onlyAdminCanCreatePipes
                    onlyAdminCanInviteUsers
                    pipes {
                        id
                        name
                        labels {
                            id
                            name
                            color
                        }
                        noun
                        organizationId
                    }
                    pipesCount
                    planName
                    role
                    users {
                        id
                        email
                        name
                        username
                    }
                    uuid
                }
            }";

        private string queryPhases = @"
            {
                pipe (id: $pipeId) {
                    phases {
                        name
                        id
                        cards_count
                        cards_can_be_moved_to_phases {
                            name
                        }
                    }
                }
            }";

        private string queryPhaseCards = @"
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
            }";

        private string queryPhaseFields = @"
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
            }";

        private string queryCard = @"
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
            }";

        private string moveCardToPhase = @"
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
                }";

        private string updateCardField = @"
                mutation {
                    updateCardField(input: { 
                        card_id: ""<<IdCard>>"",
                        field_id: ""<<IdField>>"",
                        new_value: <<NewValue>>
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
                }";

        private string createPresignedUrl = @"
                mutation {
                    createPresignedUrl(input: { organizationId: <<OrgId>>, fileName: ""<<FileName>>"" }){
                        clientMutationId
                        url
                    }
                }
            ";

        #endregion

        public Pipefy (string apiToken)
            : this(apiToken, null) { }

        public Pipefy(string apiToken, string pipeId)
            : this(apiToken, pipeId, null) { }

        public Pipefy (string apiToken, string pipeId, string uri)
        {
            Uri = uri ?? PipefyDefaultEndPoint;

            PipeId = pipeId;
            Token = apiToken;

            JsonSerializerSettingsSnake = CreateJsonSerializerSetting<SnakeCaseNamingStrategy>();
            JsonSerializerSettingsCamel = CreateJsonSerializerSetting<CamelCaseNamingStrategy>();

        }

        public static JsonSerializerSettings CreateJsonSerializerSetting<T>() where T : NamingStrategy
        {
            var contractResolver = new DefaultContractResolver
            {
                NamingStrategy = (T)Activator.CreateInstance(typeof(T)),
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
            var query = queryUserInfo;

            return GraphQl.Query<MeResult>(query, Uri, Token, JsonSerializerSettingsSnake);
        }

        public async Task<GraphQlResponse<MeResult>> QueryUserInfoAsync()
        {
            var query = queryUserInfo;

            return await GraphQl.QueryAsync<MeResult>(query, Uri, Token, JsonSerializerSettingsSnake);
        }

        public GraphQlResponse<PipeResult> QueryPipe(string pipeId, bool noPhaseInfo = false)
        {
            var selectedQuery = queryPipe;
            if (noPhaseInfo) selectedQuery = queryPipeSimple;

            var query = selectedQuery.Replace("<<PipeId>>", pipeId);

            return GraphQl.Query<PipeResult>(query, Uri, Token, JsonSerializerSettingsCamel);
        }

        public async Task<GraphQlResponse<PipeResult>> QueryPipeAsync(string pipeId, bool noPhaseInfo = false)
        {
            var selectedQuery = queryPipe;
            if (noPhaseInfo) selectedQuery = queryPipeSimple;

            var query = selectedQuery.Replace("<<PipeId>>", pipeId);

            return await GraphQl.QueryAsync<PipeResult>(query, Uri, Token, JsonSerializerSettingsCamel);
        }

        public GraphQlResponse<OrganizationsResult> QueryOrganizations()
        {
            string query = queryOrganizations;

            return GraphQl.Query<OrganizationsResult>(query, Uri, Token, CreateJsonSerializerSetting<CamelCaseNamingStrategy>());
        }

        public async Task<GraphQlResponse<OrganizationsResult>> QueryOrganizationsAsync()
        {
            string query = queryOrganizations;

            return await GraphQl.QueryAsync<OrganizationsResult>(query, Uri, Token, CreateJsonSerializerSetting<CamelCaseNamingStrategy>());
        }

        public GraphQlResponse<PipeResult> QueryPhases()
        {
            if (PipeId == null)
            {
                throw new RpaLibException("Null Pipe ID: To query info about the phases the Pipe ID is mandatory.");
            }

            return QueryPhases(PipeId);
        }
            
        public GraphQlResponse<PipeResult> QueryPhases(string pipeId)
        {
            string query = queryPhases.Replace("$pipeId", pipeId);

            return GraphQl.Query<PipeResult>(query, Uri, Token, JsonSerializerSettingsSnake);
        }

        public async Task<GraphQlResponse<PipeResult>> QueryPhasesAsync()
        {
            if (PipeId == null)
            {
                throw new RpaLibException("Null Pipe ID: To query info about the phases the Pipe ID is mandatory.");
            }

            return await QueryPhasesAsync(PipeId);
        }

        public async Task<GraphQlResponse<PipeResult>> QueryPhasesAsync(string pipeId)
        {
            string query = queryPhases.Replace("$pipeId", pipeId);

            return await GraphQl.QueryAsync<PipeResult>(query, Uri, Token, JsonSerializerSettingsSnake);
        }

        public GraphQlResponse<PhaseResult> QueryPhaseCards(string phaseId, int max = 30, string afterCursor = null)
        {
            string query = queryPhaseCards.Replace("$phaseId", phaseId).Replace("$max", max.ToString());

            if (string.IsNullOrEmpty(afterCursor))
            {
                query = query.Replace(@", after: ""$afterCursor""", string.Empty);
            }
            else
            {
                query = query.Replace("$afterCursor", afterCursor);
            }

            return GraphQl.Query<PhaseResult>(query, Uri, Token, JsonSerializerSettingsSnake);
        }

        public async Task<GraphQlResponse<PhaseResult>> QueryPhaseCardsAsync(string phaseId, int max = 30, string afterCursor = null)
        {
            string query = queryPhaseCards.Replace("$phaseId", phaseId).Replace("$max", max.ToString());

            if (string.IsNullOrEmpty(afterCursor))
            {
                query = query.Replace(@", after: ""$afterCursor""", string.Empty);
            }
            else
            {
                query = query.Replace("$afterCursor", afterCursor);
            }

            return await GraphQl.QueryAsync<PhaseResult>(query, Uri, Token, JsonSerializerSettingsSnake);
        }

        public enum OrderCardsBy
        {
            Nothing,
            Newer,
            Older
        }

        private dynamic QueryAllPhaseCardsInnerLogic(dynamic fn, string phaseId, int limit = 0, OrderCardsBy orderBy = OrderCardsBy.Nothing)
        {
            List<CardEdge> cardEdges = new List<CardEdge>();
            dynamic response;

            string pageCursor = null;
            bool hasNextPage = false;
            bool limitReached = false;

            do
            {
                response = fn(pageCursor);

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

        public GraphQlResponse<PhaseResult> QueryAllPhaseCards(string phaseId, int limit = 0, OrderCardsBy orderBy = OrderCardsBy.Nothing)
        {
            var taskQueryAll = QueryAllPhaseCardsAsync(phaseId, limit, orderBy);
            taskQueryAll.Wait();
            return taskQueryAll.Result;
        }

        /* Legacy
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
        */

        public async Task<GraphQlResponse<PhaseResult>> QueryAllPhaseCardsAsync(string phaseId, int limit = 0, OrderCardsBy orderBy = OrderCardsBy.Nothing)
        {
            List<CardEdge> cardEdges = new List<CardEdge>();
            GraphQlResponse<PhaseResult> response;

            string pageCursor = null;
            bool hasNextPage = false;
            bool limitReached = false;

            do
            {
                response = await QueryPhaseCardsAsync(phaseId, afterCursor: pageCursor);

                var phaseCards = response.Data.Phase.Cards.Edges;
                hasNextPage = response.Data.Phase.Cards.Pageinfo.Hasnextpage;

                if (hasNextPage)
                    pageCursor = response.Data.Phase.Cards.Pageinfo.Endcursor;

                //if (phaseCards.Count == 0)
                //{
                //    Trace.WriteLine($"No cards in phase \"[ID: {phaseId}]\". Total cards: {phaseCards.Count}");
                //    return null;
                //}

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
            string query = queryPhaseFields.Replace("$phaseId", phaseId);

            return GraphQl.Query<PhaseResult>(query, Uri, Token, JsonSerializerSettingsSnake);
        }

        public async Task<GraphQlResponse<PhaseResult>> QueryPhaseFieldsAsync(string phaseId)
        {
            string query = queryPhaseFields.Replace("$phaseId", phaseId);

            return await GraphQl.QueryAsync<PhaseResult>(query, Uri, Token, JsonSerializerSettingsSnake);
        }

        public GraphQlResponse<CardResult> QueryCard(string cardId)
        {
            string query = queryCard.Replace("$cardId", cardId);

            return GraphQl.Query<CardResult>(query, Uri, Token, JsonSerializerSettingsSnake);
        }

        public async Task<GraphQlResponse<CardResult>> QueryCardAsync(string cardId)
        {
            string query = queryCard.Replace("$cardId", cardId);

            return await GraphQl.QueryAsync<CardResult>(query, Uri, Token, JsonSerializerSettingsSnake);
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
            var query = moveCardToPhase.Replace("<<IdCard>>", cardId).Replace("<<IdPhase>>", destPhaseId);

            return GraphQl.Query<CardResult>(query, Uri, Token);
        }

        public async Task<GraphQlResponse<CardResult>> MoveCardToPhaseAsync(string cardId, string destPhaseId)
        {
            var query = moveCardToPhase.Replace("<<IdCard>>", cardId).Replace("<<IdPhase>>", destPhaseId);

            return await GraphQl.QueryAsync<CardResult>(query, Uri, Token);
        }

        private string GetListForUpdateCardField(string[] newValues)
        {
            List<string> escapedNewValues = new List<string>();
            string newValueStringList;

            foreach (var newValue in newValues)
            {
                var escaped = newValue.Replace(@"\", @"\\").Replace(@"""", @"\""");
                var quotedAfterEscaped = $"\"{escaped}\"";
                escapedNewValues.Add(quotedAfterEscaped);
            }

            var newValuesJoined = string.Join("\", ", escapedNewValues);
            newValueStringList = $"[{newValuesJoined}]";

            return newValueStringList;
        }

        public GraphQlResponse<UpdateCardFieldResult> UpdateCardField(string cardId, string fieldId, string[] newValues)
        {
            return UpdateCardField(cardId, fieldId, GetListForUpdateCardField(newValues));
        }

        public GraphQlResponse<UpdateCardFieldResult> UpdateCardField(string cardId, string fieldId, string newValue)
        {
            var treatedValue = !Ut.IsMatch(newValue, @"\[[^\]]+\]") ? $"\"{newValue}\"" : newValue;
            var query = updateCardField.Replace("<<IdCard>>", cardId).Replace("<<IdField>>", fieldId).Replace("<<NewValue>>", treatedValue);

            return GraphQl.Query<UpdateCardFieldResult>(query, Uri, Token);
        }

        public async Task<GraphQlResponse<UpdateCardFieldResult>> UpdateCardFieldAsync(string cardId, string fieldId, string[] newValues)
        {
            return await UpdateCardFieldAsync(cardId, fieldId, GetListForUpdateCardField(newValues));
        }

        public async Task<GraphQlResponse<UpdateCardFieldResult>> UpdateCardFieldAsync(string cardId, string fieldId, string newValue)
        {
            var treatedValue = !Ut.IsMatch(newValue, @"\[[^\]]+\]") ? $"\"{newValue}\"" : newValue;
            var query = updateCardField.Replace("<<IdCard>>", cardId).Replace("<<IdField>>", fieldId).Replace("<<NewValue>>", treatedValue);

            return await GraphQl.QueryAsync<UpdateCardFieldResult>(query, Uri, Token);
        }

        public GraphQlResponse<CreatePresignedUrlResult> CreatePresignedUrl(string fileBaseName)
        {
            if (PipeId == null)
            {
                throw new RpaLibException("Null Pipe ID: To create a presigned URL the Pipe ID is mandatory.");
            }

            var organizationId = QueryPipe(PipeId).Data.Pipe.OrganizationId;

            var query = createPresignedUrl.Replace("<<OrgId>>", organizationId).Replace("<<FileName>>", fileBaseName);

            return GraphQl.Query<CreatePresignedUrlResult>(query, Uri, Token, JsonSerializerSettingsCamel);
        }

        public async Task<GraphQlResponse<CreatePresignedUrlResult>> CreatePresignedUrlAsync(string fileBaseName)
        {
            if (PipeId == null)
            {
                throw new RpaLibException("Null Pipe ID: To create a presigned URL the Pipe ID is mandatory.");
            }

            var organizationId = (await QueryPipeAsync(PipeId)).Data.Pipe.OrganizationId;

            var query = createPresignedUrl.Replace("<<OrgId>>", organizationId).Replace("<<FileName>>", fileBaseName);

            return await GraphQl.QueryAsync<CreatePresignedUrlResult>(query, Uri, Token, JsonSerializerSettingsCamel);
        }

        #endregion

        /// <summary>
        /// Query pipefy info using API calls and redirects the info output to the Tracing.
        /// </summary>
        /// <param name="infoType">A PipefyInfo enum type indicating the kind of information to retrieve.</param>
        /// <param name="phaseId">The Pipefy phase ID that is needed to perform some queries.</param>
        /// <exception cref="ArgumentNullException">Thrown when a parameter is needed by some query but wasn't provided.</exception>
        public async Task<string> ShowInfoAsync(PipefyInfo infoType, string phaseId = null, string cardId = null)
        {
            StringBuilder infoMsg = new StringBuilder();

            switch (infoType)
            {
                case PipefyInfo.PhasesAndCardsCount:
                    var phases = (await QueryPhasesAsync()).Data.Pipe.Phases;

                    infoMsg.AppendLine("Phases and available cards: ");
                    foreach (Phase phase in phases)
                    {
                        var canBeMovedTo = string.Join(", ", phase.CardsCanBeMovedToPhases.Select(x => $"\"{x.Name}\"").ToArray());
                        var msg = string.Join(Environment.NewLine,
                            $"- Phase Name: \"{phase.Name}\"",
                            $"  Phase ID: {phase.Id}",
                            $"  Number of cards in phase: {phase.CardsCount}",
                            $"  Cards can be moved to phases: {canBeMovedTo}");
                        infoMsg.AppendLine(msg + Environment.NewLine);
                    }
                    break;

                case PipefyInfo.PhaseFields:
                    if (phaseId == null)
                        throw new RpaLibArgumentNullException("Phase ID argument is mandatory to fetch the fields from that phase.");
                    PhaseResult phaseQuery = (await QueryPhaseFieldsAsync(phaseId)).Data;
                    infoMsg.AppendLine($"Phase (ID {phaseQuery.Phase.Id}) {phaseQuery.Phase.Name}, has the following fields:");
                    foreach (PhaseField field in phaseQuery.Phase.Fields)
                    {
                        string msg = string.Join("\n",
                            "- ID: " + field.Id,
                            "  Label: " + field.Label,
                            "  Options: " + string.Join(", ", field.Options));
                        infoMsg.AppendLine(msg);
                    }
                    break;

                case PipefyInfo.CardFields:
                    if (cardId == null)
                        throw new RpaLibArgumentNullException("Card ID argument is mandatory to fetch the fields from that card.");
                    CardResult cardQuery = (await QueryCardAsync(cardId)).Data;
                    infoMsg.AppendLine($"Card (ID {cardQuery.Card.Id}) has the following fields");
                    foreach (CardField field in cardQuery.Card.Fields)
                    {
                        string msg = string.Join(Environment.NewLine,
                            "- ID: " + field.Field.Id,
                            "  Name: " + field.Name,
                            "  Options: " + string.Join(", ", field.Field.Options));
                        infoMsg.AppendLine(msg);
                    }
                    break;

                case PipefyInfo.Organizations:
                    OrganizationsResult organizationsQuery = (await QueryOrganizationsAsync()).Data;
                    infoMsg.AppendLine($"Organizations registered within this accound are:");

                    var getPipesString = new Func<Organization, string>((org) =>
                    {
                        StringBuilder pipesStr = new StringBuilder();

                        foreach (var pipe in org.Pipes)
                        {
                            var currentPipeInfo = string.Join(Environment.NewLine,
                            "  - Pipe ID: " + pipe.Id,
                            "    Pipe Name: " + pipe.Name);

                            pipesStr.AppendLine(currentPipeInfo);
                        }

                        return pipesStr.ToString();
                    });

                    var getUsersString = new Func<Organization, string>((org) =>
                    {
                        StringBuilder usrStr = new StringBuilder();

                        foreach (var user in org.Users)
                        {
                            var currentPipeInfo = string.Join(Environment.NewLine,
                            "  - User ID: " + user.Id,
                            "    User Name: " + user.Name,
                            "    User Email: " + user.Email,
                            "    User Username: " + user.Username);

                            usrStr.AppendLine(currentPipeInfo);
                        }

                        return usrStr.ToString();
                    });

                    foreach (var org in organizationsQuery.Organizations)
                    {
                        string msg = string.Join(Environment.NewLine,
                            "- Org ID: " + org.Id,
                            "  Org Name: " + org.Name,
                            "  PipesCount: " + org.PipesCount,
                            "  Pipes: ",
                            getPipesString(org),
                            "  Users:",
                            getUsersString(org),
                            "  PlanName: " + org.PlanName,
                            "  Role: " + org.Role
                            );
                        infoMsg.AppendLine(msg);
                    }
                    break;

                case PipefyInfo.CurrentUser:
                    var meQuery = (await QueryUserInfoAsync()).Data;
                    infoMsg.AppendLine($"Current connected user:");
                    infoMsg.AppendLine($"  ID: {meQuery.Me.Id}");
                    infoMsg.AppendLine($"  Name: {meQuery.Me.Name}");
                    infoMsg.AppendLine($"  Email: {meQuery.Me.Email}");
                    infoMsg.AppendLine($"  Username: {meQuery.Me.Username}");

                    break;
            }

            return infoMsg.ToString();
        }

        public void ShowInfo(PipefyInfo infoType, string phaseId = null, string cardId = null)
        {
            Task<string> t1 = ShowInfoAsync(infoType, phaseId, cardId);

            t1.Wait();

            Trace.WriteLine(t1.Result, withTimeSpec: false, color: ConsoleColor.Yellow);
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

        public string AttachFileToCard(string filePath, string cardId, string fieldId)
        {
            var fileBaseName = Path.GetFileName(filePath);
            var url = CreatePresignedUrl(fileBaseName).Data.CreatePresignedUrl.Url;

            var resp1 = Ut.Curl(request: "PUT", url: url, header: new string[] { "Content-Type: application/excel" }, data: "BINARY_DATA");
            var resp2 = Ut.Curl(uploadFilePath: filePath, url: url);

            var pathFromUrl = Ut.Replace(url, @"^https://[^/]+/", string.Empty);
            pathFromUrl = Ut.Match(pathFromUrl, @"^.+" + fileBaseName.Replace(".", @"\."));
            var result = UpdateCardField(cardId, fieldId, new string[] { pathFromUrl });

            return null;
        }
    }
}
