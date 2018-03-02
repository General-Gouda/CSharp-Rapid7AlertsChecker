using RestSharp;
using Newtonsoft.Json;

namespace Rapid7AlertChecker
{
    class ApiCallers
    {
        public static dynamic Get_API_Results(string graph_api_endpoint, string resource, string token)
        {
            var client = new RestClient(graph_api_endpoint);

            var request = new RestRequest(resource, Method.GET);
            request.AddHeader("Authorization", string.Format("Bearer {0}", token));
            request.AddHeader("Accept", "application/json");
            request.AddHeader("Content-Type", "application/json");

            var result = client.Execute(request);

            dynamic content = JsonConvert.DeserializeObject(result.Content);

            dynamic values = content["value"];

            return values;
        }

        public static dynamic API_Action_Caller(string graph_api_endpoint, string resource, string token, string actionType, object json)
        {
            var client = new RestClient(graph_api_endpoint);
            var jsonBody = JsonConvert.SerializeObject(json, Formatting.None);

            RestRequest request = new RestRequest
            {
                Resource = resource,
                RequestFormat = DataFormat.Json
            };

            if (actionType.ToLower() == "post")
            {
                request.Method = Method.POST;
            }
            else if (actionType.ToLower() == "patch")
            {
                request.Method = Method.PATCH;
            }

            request.AddHeader("Authorization", string.Format("Bearer {0}", token));
            request.AddHeader("Accept", "application/json");
            request.AddHeader("Content-Type", "application/json");
            request.AddParameter("application/json", jsonBody, ParameterType.RequestBody);

            var result = client.Execute(request);

            dynamic content = JsonConvert.DeserializeObject(result.Content);

            return content;
        }

        public static bool MoveEmail(string archiveMailFolderID, string graph_api_endpoint, string username, string emailID, string token)
        {
            var moveEmailBody = new MoveEmailBody
            {
                DestinationId = archiveMailFolderID
            };

            try
            {
                ApiCallers.API_Action_Caller(
                    graph_api_endpoint,
                    string.Format("/users/{0}/messages/{1}/move", username, emailID),
                    token,
                    "post",
                    moveEmailBody
                );

                return true;
            }
            catch (System.Exception)
            {
                return false;
            }
        }

        public static bool MarkAsRead(string graph_api_endpoint, string username, string emailID, string token)
        {
            var markAsReadBody = new JsonBody
            {
                IsRead = true
            };

            try
            {
                ApiCallers.API_Action_Caller(
                    graph_api_endpoint,
                    string.Format("/users/{0}/messages/{1}", username, emailID),
                    token,
                    "patch",
                    markAsReadBody
                );

                return true;
            }
            catch (System.Exception)
            {
                return false;
            }
        }
    }
}
