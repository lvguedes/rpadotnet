using RpaLib.APIs;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Http;
using RpaLib.APIs.GraphQL;
using RpaLib.APIs.Pipefy.Model;
using RpaLib.APIs.Pipefy.Exception;
using RpaLib.Tracing;
using RpaLib.Exceptions;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
using System.IO;
using RpaLib.ProcessAutomation;
using System.Security.Policy;

namespace RpaLib.APIs.Pipefy
{
    public class Pipefy
    {
        private string _msgPipeIdRequired = @"Pipe ID must be given to the object constructor or as parameter to this method.";
        private List<PhaseField> _startFormFields;
        private List<PipeReport> _pipeReports;
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

        private string queryPipeReports = @"
            query {
                pipe(id: ""<<PipeId>>"") {
                    reports {
                        id
                        name
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
                        fields {
                            id
                            options
                            label
                            type 
                        }
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
                fields {
                  label
                  id
                  options
                  type
                  is_multiple
                  description
                  editable
                  required
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

        private string queryStartFormFields = @"
            query {
                pipe(id: ""<<PipeId>>"") {
                    start_form_fields {
                        id
                        options
                        label
                        type
                    }
                }
            }
            ";

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
                        downloadUrl
                    }
                }
            ";

        private string cardsImporter = @"
	          mutation { 
                cardsImporter(
                  input: <<ParseInput>> 
                ) { 
                  cardsImportation { 
                    id
                    status
                    importedCards
                  } 
                } 
              }
            ";

        private string exportReportCreateObj = @"
            mutation {
                exportPipeReport (input: {pipeId: ""<<PipeId>>"", pipeReportId: ""<<PipeReportId>>""}) {
                    pipeReportExport {
                        id
                    }
                }
            }
        ";

        private string exportReportExportGetUrl = @"
            query {
                pipeReportExport(id: ""<<ExportObjId>>"") {
                    fileURL
                    state
                    startedAt
                    requestedBy {
                        id
                    }
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

        public GraphQlResponse<PipeResult> QueryPipe(bool noPhaseInfo = false)
        {
            if (PipeId == null)
                throw new RpaLibException(_msgPipeIdRequired);
            else
                return QueryPipe(PipeId, noPhaseInfo);
        }

        public GraphQlResponse<PipeResult> QueryPipe(string pipeId, bool noPhaseInfo = false)
        {
            var selectedQuery = queryPipe;
            if (noPhaseInfo) selectedQuery = queryPipeSimple;

            var query = selectedQuery.Replace("<<PipeId>>", pipeId);

            return GraphQl.Query<PipeResult>(query, Uri, Token, JsonSerializerSettingsCamel);
        }

        public async Task<GraphQlResponse<PipeResult>> QueryPipeAsync(bool noPhaseInfo = false)
        {
            if (PipeId == null)
                throw new RpaLibException(_msgPipeIdRequired);
            else
                return await QueryPipeAsync(PipeId, noPhaseInfo);
        }

        public async Task<GraphQlResponse<PipeResult>> QueryPipeAsync(string pipeId, bool noPhaseInfo = false)
        {
            var selectedQuery = queryPipe;
            if (noPhaseInfo) selectedQuery = queryPipeSimple;

            var query = selectedQuery.Replace("<<PipeId>>", pipeId);

            return await GraphQl.QueryAsync<PipeResult>(query, Uri, Token, JsonSerializerSettingsCamel);
        }

        public GraphQlResponse<PipeResult> QueryPipeReports()
        {
            if (PipeId == null)
                throw new RpaLibException(_msgPipeIdRequired);
            else
                return QueryPipeReports(PipeId);
        }

        public GraphQlResponse<PipeResult> QueryPipeReports(string pipeId)
        {
            var query = queryPipeReports.Replace("<<PipeId>>", pipeId);

            return GraphQl.Query<PipeResult>(query, Uri, Token, JsonSerializerSettingsCamel);
        }

        public async Task<GraphQlResponse<PipeResult>> QueryPipeReportsAsync()
        {
            if (PipeId == null)
                throw new RpaLibException(_msgPipeIdRequired);
            else
                return await QueryPipeReportsAsync(PipeId);
        }

        public async Task<GraphQlResponse<PipeResult>> QueryPipeReportsAsync(string pipeId)
        {
            var query = queryPipeReports.Replace("<<PipeId>>", pipeId);

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
                throw new RpaLibException(_msgPipeIdRequired);
            else
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
                throw new RpaLibException(_msgPipeIdRequired);
            else
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

        public GraphQlResponse<PipeResult> QueryStartFormFields()
        {
            string query = queryStartFormFields.Replace(@"<<PipeId>>", PipeId);

            return GraphQl.Query<PipeResult>(query, Uri, Token, JsonSerializerSettingsSnake);
        }

        public async Task<GraphQlResponse<PipeResult>> QueryStartFormFieldsAsync()
        {
            string query = queryStartFormFields.Replace(@"<<PipeId>>", PipeId);

            return await GraphQl.QueryAsync<PipeResult>(query, Uri, Token, JsonSerializerSettingsSnake);
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

        private string CardsImporterParseInput(string url, ExcelColumn assigneesColumn = ExcelColumn.None, ExcelColumn labelsColumn = ExcelColumn.None,
            ExcelColumn dueDateColumn = ExcelColumn.None, ExcelColumn currentPhaseColumn = ExcelColumn.None, Dictionary<ExcelColumn, string> fieldValuesColumn = null)
        {
            StringBuilder input = new StringBuilder($"{{ pipeId: \"{PipeId}\", url: \"{url}\"");

            if (assigneesColumn != ExcelColumn.None)
                input.Append($", assigneesColumn: \"{assigneesColumn}\"");

            if (labelsColumn != ExcelColumn.None)
                input.Append($", labelsColumn: \"{labelsColumn}\"");

            if (dueDateColumn != ExcelColumn.None)
                input.Append($", dueDateColumn: \"{dueDateColumn}\"");

            if (currentPhaseColumn != ExcelColumn.None)
                input.Append($", currentPhaseColumn: \"{currentPhaseColumn}\"");

            if (fieldValuesColumn != null)
            {
                input.Append($", fieldValuesColumns: [ ");
                var processedItems = 0;
                foreach (var columnField in fieldValuesColumn)
                {
                    if (processedItems > 0)
                        input.Append(", "); // item separator from second to the end

                    input.Append($"{{ column: \"{columnField.Key}\", fieldId: \"{columnField.Value}\" }}");
                    processedItems++;
                }
                input.Append("] ");
            }

            input.Append("} ");

            return input.ToString();
        }

        public async Task<GraphQlResponse<CardsImporterResult>> CardsImporterAsync(string excelFilePath, ExcelColumn assigneesColumn = ExcelColumn.None,
            ExcelColumn labelsColumn = ExcelColumn.None, ExcelColumn dueDateColumn = ExcelColumn.None, ExcelColumn currentPhaseColumn = ExcelColumn.None,
            Dictionary<ExcelColumn, string> exlColumnPipField = null)
        {
            var cloudFileName = Path.GetFileName(excelFilePath);
            var uploadResult = await UploadFileAsync(excelFilePath, cloudFileName);

            var input = CardsImporterParseInput(uploadResult.DownloadUrl, assigneesColumn, labelsColumn, dueDateColumn, currentPhaseColumn, exlColumnPipField);

            /* Exemplo de como deve ficar a query:
             mutation { 
                cardsImporter(
                  input: { 
                    pipeId: "902791",
                    url: "https://app.pipefy.com/storage/v1/signed/orgs/058fec5c-828c-4220-9f44-ccb4e183f8a8/uploads/c8b5f927-8dab-468f-9aa2-c491aac5759c/Cronograma_de_Fechamento_Import_Pipefy_maio.2024.xlsx?signature=eRta4CngLFDvuuT7%2Fr6qbop3Dny3yRG%2F2BeUJZea%2F0M%3D",
                    fieldValuesColumns: [ { column: "B", fieldId: "solciita_o" },
          										            { column: "C", fieldId: "detalhes" },
          										            { column: "D", fieldId: "prazo_ideal" },
          										            { column: "E", fieldId: "frequency" },
          										            { column: "F", fieldId: "action_plan_performer" },
          										            { column: "G", fieldId: "performer_email" }, 
          										            { column: "H", fieldId: "entidade_respons_vel" },
          										            { column: "I", fieldId: "fiscal_year_period" },
          										            { column: "J", fieldId: "copy_of_per_odo_do_ano_fiscal" },
          										            { column: "K", fieldId: "aprovador_respons_vel_1" },
          										            { column: "L", fieldId: "approver_email_address" } ] 
      
                  }) { 
                  cardsImportation { 
                    id
                    status
                    importedCards
                  } 
                } 
              }
             */
            var query = cardsImporter.Replace("<<ParseInput>>", input);

            return await GraphQl.QueryAsync<CardsImporterResult>(query, Uri, Token, JsonSerializerSettingsCamel);
        }

        public GraphQlResponse<CardsImporterResult> CardsImporter(string excelFilePath, ExcelColumn assigneesColumn = ExcelColumn.None,
            ExcelColumn labelsColumn = ExcelColumn.None, ExcelColumn dueDateColumn = ExcelColumn.None, ExcelColumn currentPhaseColumn = ExcelColumn.None, 
            Dictionary<ExcelColumn, string> exlColumnPipField = null)
        {
            var asyncVersion = CardsImporterAsync(excelFilePath, assigneesColumn, labelsColumn, dueDateColumn, currentPhaseColumn, exlColumnPipField);
            asyncVersion.Wait();
            return asyncVersion.Result;
        }

        #endregion

        #region DisplayInfo

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
                    var phaseQueryResp = await QueryPhaseFieldsAsync(phaseId);
                    PhaseResult phaseQuery = phaseQueryResp.Data;
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

        #endregion

        #region Utils

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

        public string AttachFileToCardCurl(string filePath, string cardId, string fieldId)
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

        public GraphQlResponse<UpdateCardFieldResult> AttachFileToCard(string filePath, string cardId, string fieldId)
        {
            var asyncVersion = AttachFileToCardAsync(filePath, cardId, fieldId);
            asyncVersion.Wait();
            return asyncVersion.Result;
        }

        public async Task<GraphQlResponse<UpdateCardFieldResult>> AttachFileToCardAsync(string filePath, string cardId, string fieldId)
        {
            var fileBaseName = Path.GetFileName(filePath);
            var uploadFileReturn = await UploadFileAsync(filePath, fileBaseName);

            var pathFromUrl = Ut.Replace(uploadFileReturn.UploadUrl, @"^https://[^/]+/", string.Empty);
            pathFromUrl = Ut.Match(pathFromUrl, @"^.+" + fileBaseName.Replace(".", @"\."));

            var result = await UpdateCardFieldAsync(cardId, fieldId, new string[] { pathFromUrl });

            return result;
        }

        public UploadFileReturn UploadFile(string filePath)
        {
            var fileNameInCloudWithExtension = Path.GetFileName(filePath);
            return UploadFile(filePath, fileNameInCloudWithExtension);
        }

        public UploadFileReturn UploadFile(string filePath, string fileNameInCloudWithExtension)
        {
            var asyncVersion = UploadFileAsync(filePath, fileNameInCloudWithExtension);

            asyncVersion.Wait();

            return asyncVersion.Result;
        }


        public async Task<UploadFileReturn> UploadFileAsync(string filePath)
        {
            var fileNameInCloudWithExtension = Path.GetFileName(filePath);
            return await UploadFileAsync(filePath, fileNameInCloudWithExtension);
        }

        public async Task<UploadFileReturn> UploadFileAsync(string filePath, string fileNameInCloudWithExtension)
        {
            var presignedUrl = (await CreatePresignedUrlAsync(fileNameInCloudWithExtension))?.Data.CreatePresignedUrl;
            var uploadUrl = presignedUrl.Url;
            var downloadUrl = presignedUrl.DownloadUrl;

            var testFilePath = Ut.GetFullPath(filePath);

            var response = await Ut.HttpPutFileAsync(testFilePath, uploadUrl);

            if (!response.IsSuccessStatusCode)
            {
                throw new RpaLibException($"The file \"{filePath}\" could not be uploaded to AWS bucket (via presigned URL)");
            }

            return new UploadFileReturn(uploadUrl, downloadUrl);
        }

        private string MsgApiRespNotSuccess<T>(GraphQlResponse<T> response)
        {
            return $"Status code is not success for querying start form fields through" +
                        $" query: {response.Queried}" + Environment.NewLine +
                        $"Status code: {response.StatusCode}";
        }

        public string GetReportId(string reportNameRegex, bool queryOnlyOnFirstCall = false)
        {
            var asyncVersion = GetReportIdAsync(reportNameRegex, queryOnlyOnFirstCall);
            asyncVersion.Wait();
            return asyncVersion.Result;
        }

        public async Task<string> GetReportIdAsync(string reportNameRegex, bool queryOnlyOnFirstCall = false)
        {
            if (queryOnlyOnFirstCall == false || (queryOnlyOnFirstCall == true && _pipeReports == null))
            {
                var resp = await QueryPipeReportsAsync();

                if (!resp.IsSuccessStatusCode)
                    throw new RpaLibException(MsgApiRespNotSuccess(resp));
                else
                    _pipeReports = resp.Data.Pipe.Reports;
            }

            var found = _pipeReports.Where(x => Ut.IsMatch(x.Name, reportNameRegex));

            if (found.Count() > 1)
            {
                var foundCsv = string.Join(", ", found.Select(x => $"{x.Name}[{x.Id}]"));
                throw new RpaLibException($"More than one Pipe Report match the regex \"{reportNameRegex}\": {foundCsv}");
            }
            else if (found.Count() == 0)
            {
                throw new RpaLibException($"No Pipe Reports found in which their name that match the regex \"{reportNameRegex}\"");
            }
            else
                return found.ToArray()[0].Id;
        }

        public string GetStartFormFieldId(string startFormFieldLabelRegex, bool queryOnlyOnFirstCall = false)
        {
            var asyncVersion = GetStartFormFieldIdAsync(startFormFieldLabelRegex, queryOnlyOnFirstCall);
            asyncVersion.Wait();
            return asyncVersion.Result;
        }

        public async Task<string> GetStartFormFieldIdAsync(string startFormFieldLabelRegex, bool queryOnlyOnFirstCall = false)
        {
            if (queryOnlyOnFirstCall == false || ( queryOnlyOnFirstCall == true && _startFormFields == null ))
            {
                var resp = await QueryStartFormFieldsAsync();

                if (!resp.IsSuccessStatusCode)
                    throw new RpaLibException(MsgApiRespNotSuccess(resp));
                else
                    _startFormFields = resp.Data.Pipe.StartFormFields;

            }

            var found = _startFormFields.Where(x => Ut.IsMatch(x.Label, startFormFieldLabelRegex));

            if (found.Count() > 1)
            {
                var foundCsv = string.Join(", ", found.Select(x => $"{x.Label}[{x.Id}]"));
                throw new RpaLibException($"More than one Start Form Field match the regex \"{startFormFieldLabelRegex}\": {foundCsv}");
            }
            else if (found.Count() == 0)
            {
                throw new RpaLibException($"No Start Form Fields found in which their name that match the regex \"{startFormFieldLabelRegex}\"");
            }
            else
                return found.ToArray()[0].Id;
        }

        public string ExportPipeReport(string reportNameRegex, string filePath)
        {
            if (PipeId == null)
                throw new RpaLibException(_msgPipeIdRequired);
            else
                return ExportPipeReport(PipeId, reportNameRegex, filePath);
        }

        public string ExportPipeReport(string pipeId, string reportNameRegex, string filePath)
        {
            var asyncVersion = ExportPipeReportAsync(pipeId, reportNameRegex, filePath);
            asyncVersion.Wait();
            return asyncVersion.Result;
        }

        public async Task<string> ExportPipeReportAsync(string reportNameRegex, string filePath)
        {
            if (PipeId == null)
                throw new RpaLibException(_msgPipeIdRequired);
            else
                return await ExportPipeReportAsync(PipeId, reportNameRegex, filePath);
        }

        // mutation to create the export obj using report ID and Pipe ID
        private async Task<string> ExportPipeReportCreateObj(string pipeId, string reportId)
        {
            var queryExportObj = exportReportCreateObj.Replace("<<PipeId>>", pipeId).Replace("<<PipeReportId>>", reportId);
            
            var respExportObj = await GraphQl.QueryAsync<ExportPipeReportResult>(queryExportObj, Uri, Token, JsonSerializerSettingsCamel);

            if (!respExportObj.IsSuccessStatusCode)
                throw new RpaLibException(MsgApiRespNotSuccess(respExportObj));

            return respExportObj.Data.ExportPipeReport.PipeReportExport.Id;
        }

        // query to get the report download URL using the export obj ID from the mutation above
        private async Task<string> ExportPipeReportGetUrl(string exportObjId)
        {
            GraphQlResponse<PipeReportExportResult> respGetUrl;
            var queryGetUrl = exportReportExportGetUrl.Replace("<<ExportObjId>>", exportObjId);

            do
            {
                respGetUrl = await GraphQl.QueryAsync<PipeReportExportResult>(queryGetUrl, Uri, Token, JsonSerializerSettingsCamel);

                if (!respGetUrl.IsSuccessStatusCode)
                    throw new RpaLibException(MsgApiRespNotSuccess(respGetUrl));

            } while (respGetUrl.Data.PipeReportExport.State == ExpirationState.Processing);

            return respGetUrl.Data.PipeReportExport.FileURL;
        }

        public async Task<string> ExportPipeReportAsync(string pipeId, string reportNameRegex, string filePath)
        {
            // Gets the report ID
            var reportId = await GetReportIdAsync(reportNameRegex);

            var exportObjId = await ExportPipeReportCreateObj(pipeId, reportId);

            var downloadUrl = await ExportPipeReportGetUrl(exportObjId);

            Ut.DownloadFile(downloadUrl, filePath);

            return downloadUrl;
        }

        #endregion
    }
}
