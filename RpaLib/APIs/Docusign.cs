using DocuSign.CodeExamples.Authentication;
using DocuSign.CodeExamples.Common;
using DocuSign.eSign.Api;
using DocuSign.eSign.Client;
using DocuSign.eSign.Client.Auth;
using DocuSign.eSign.Model;
using UsrInfo = DocuSign.eSign.Client.Auth.OAuth.UserInfo;
using Microsoft.IdentityModel.Protocols;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using static DocuSign.eSign.Client.Auth.OAuth;
using File = System.IO.File;

namespace RpaLib.APIs
{
    public enum EnvelopeStatus
    {
        sent,
        created,
    }

    public class Docusign
    {
        private static readonly string DevCenterPage = "https://developers.docusign.com/platform/auth/consent";

        private OAuthToken AccessToken { get; set; }
        private UsrInfo.Account Account { get; set; }
        public string BasePath { get; private set; }
        public DocuSignClient Client { get; private set; }
        public EnvelopesApi EnvelopesApi { get; private set; }

        public Docusign(string clientId, string impersonatedUserId, string authServer, string privateKeyFile)
        {
            AccessToken = Authenticate(clientId, impersonatedUserId, authServer, privateKeyFile);
            Account = GetAccount(authServer, AccessToken.access_token);
            BasePath = Account.BaseUri + "/restapi";
            Client = GetDocuSignClient(BasePath, AccessToken.access_token);
            EnvelopesApi = new EnvelopesApi(Client);
        }

        public static OAuthToken Authenticate(string clientId, string impersonatedUserId, string authServer, string privateKeyFile)
        {
            Console.ForegroundColor = ConsoleColor.White;

            var accessToken = new OAuthToken();

            try
            {
                accessToken = JWTAuth.AuthenticateWithJWT("ESignature", clientId, impersonatedUserId,
                                                            authServer,DSHelper.ReadFileContent(privateKeyFile));
            }
            catch (ApiException apiExp)
            {
                // Consent for impersonation must be obtained to use JWT Grant
                if (apiExp.Message.Contains("consent_required"))
                {
                    // Caret needed for escaping & in windows URL
                    string caret = "";
                    if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                    {
                        caret = "^";
                    }

                    // build a URL to provide consent for this Integration Key and this userId
                    string url = "https://" + authServer + "/oauth/auth?response_type=code" + caret + "&scope=impersonation%20signature" + caret +
                        "&client_id=" + clientId + caret + "&redirect_uri=" + DevCenterPage;
                    Console.WriteLine($"Consent is required - launching browser (URL is {url.Replace(caret, "")})");

                    // Start new browser window for login and consent to this app by DocuSign user
                    if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                    {
                        Process.Start(new ProcessStartInfo("cmd", $"/c start {url}") { CreateNoWindow = false });
                    }
                    else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
                    {
                        Process.Start("xdg-open", url);
                    }
                    else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
                    {
                        Process.Start("open", url);
                    }

                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Unable to send envelope; Exiting. Please rerun the console app once consent was provided");
                    Console.ForegroundColor = ConsoleColor.White;
                    Environment.Exit(-1);
                }
            }
            return accessToken;
        }

        public static UsrInfo.Account GetAccount(string authServer, string accessToken)
        {
            var docuSignClient = new DocuSignClient();
            docuSignClient.SetOAuthBasePath(authServer);
            
            UsrInfo userInfo = docuSignClient.GetUserInfo(accessToken);
            
            return userInfo.Accounts.FirstOrDefault();
        }

        public static DocuSignClient GetDocuSignClient(string basePath, string accessToken)
        {
            var docuSignClient = new DocuSignClient(basePath);
            docuSignClient.Configuration.DefaultHeader.Add("Authorization", "Bearer " + accessToken);
            return docuSignClient;
        }

        public EnvelopeAuditEventResponse GetEnvelopeHistory(string envelopeId)
        {
            return EnvelopesApi.ListAuditEvents(Account.AccountId, envelopeId);
        }

        public static EnvelopeDefinition MakeEnvelope(Document[] documents, Signer[] signers, CarbonCopy[] carbonCopies,
            EnvelopeStatus envStatus = EnvelopeStatus.sent, string emailSubject = "Please, sign the document(s).")
        {
            EnvelopeDefinition envelope = new EnvelopeDefinition();

            envelope.EmailSubject = emailSubject;
            envelope.Documents = documents.ToList();

            Recipients recipients = new Recipients
            {
                Signers = signers.ToList(),
                CarbonCopies = carbonCopies.ToList(),
            };
            envelope.Recipients = recipients;

            envelope.Status = envStatus.ToString();

            return envelope;
        }

        public static EnvelopeDefinition ExampleOfEnvelope()
        {
            string docPath = @"C:\Users\lsilva46\OneDrive - Capgemini\Documents\POC DocuSign\Document to Sign.pdf";

            string docPdfBytes = Convert.ToBase64String(File.ReadAllBytes(docPath));

            // To make an envelope we need these 3 information
            Signer[] signers;
            CarbonCopy[] carbonCopies;
            Document[] documents;

            Document document = new Document
            {
                DocumentBase64 = docPdfBytes,
                Name = "PDF doc to sign",
                FileExtension = "pdf",
                DocumentId = "1",
            };

            // Signer definition
            Signer company = new Signer
            {
                Email = "lucasvguedess@gmail.com",
                Name = "Company",
                RecipientId = "1",
                RoutingOrder = "1",
            };
            Signer employee = new Signer
            {
                Email = "lucas.n.silva@capgemini.com",
                Name = "Mr. Employee",
                RecipientId = "2",
                RoutingOrder = "2",
            };

            // Carbon copy definitions
            CarbonCopy companycc = new CarbonCopy
            {
                Email = "lucasvguedess@gmail.com",
                Name = "Company",
                RecipientId = "3",
                RoutingOrder = "3",
            };

            // Positions where signatures must be placed
            SignHere firstSignPosition = new SignHere
            {
                PageNumber = "1",
                DocumentId = "1",
                XPosition = "144",
                YPosition = "672",
            };
            SignHere secondSignPosition = new SignHere
            {
                PageNumber = "1",
                DocumentId = "1",
                XPosition = "144",
                YPosition = "702",
            };

            /*
            document.Tabs = new Tabs
            {
                SignHereTabs = new List<SignHere> { firstSignPosition },
            };
            */

            // who must sign in which position? we define below
            employee.Tabs = new Tabs
            {
                SignHereTabs = new List<SignHere> { firstSignPosition },
            };
            company.Tabs = new Tabs
            {
                SignHereTabs = new List<SignHere> { secondSignPosition },
            };

            // grouping the info defined above
            documents = new Document[] { document };
            signers = new Signer[] { employee, company };
            carbonCopies = new CarbonCopy[] { companycc };

            // passing grouped info and creating an envelope
            EnvelopeDefinition myEnvelope = MakeEnvelope(documents, signers, carbonCopies);

            return myEnvelope;
        }

        public EnvelopeSummary SendEnvelopeViaEmail(EnvelopeDefinition envelopeDefinition)
        {
            return EnvelopesApi.CreateEnvelope(Account.AccountId, envelopeDefinition);
        }

        public EnvelopeUpdateSummary ResendEnvelope(string envelopeId)
        {
            var options = new EnvelopesApi.UpdateOptions();
            options.resendEnvelope = "true";
            var envelope = EnvelopesApi.GetEnvelope(Account.AccountId, envelopeId);
            return EnvelopesApi.Update(Account.AccountId, envelopeId, new Envelope { EnvelopeId = envelopeId }, options);
        }
    }
}
