// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license.
//
// Microsoft Bot Framework: http://botframework.com
// Microsoft Teams: https://dev.office.com/microsoft-teams
//
// Bot Builder SDK GitHub:
// https://github.com/Microsoft/BotBuilder
//
// Bot Builder SDK Extensions for Teams
// https://github.com/OfficeDev/BotBuilder-MicrosoftTeams
//
// Copyright (c) Microsoft Corporation
// All rights reserved.
//
// MIT License:
// Permission is hereby granted, free of charge, to any person obtaining
// a copy of this software and associated documentation files (the
// "Software"), to deal in the Software without restriction, including
// without limitation the rights to use, copy, modify, merge, publish,
// distribute, sublicense, and/or sell copies of the Software, and to
// permit persons to whom the Software is furnished to do so, subject to
// the following conditions:
//
// The above copyright notice and this permission notice shall be
// included in all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED ""AS IS"", WITHOUT WARRANTY OF ANY KIND,
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Connector;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using TeamsAdmin.Models;

namespace TeamsAdmin.Helper
{
    /// <summary>
    /// Provides all the functionality for Microsft Teams Graph APIs
    /// </summary>
    public class GraphAPIHelper
    {
        static readonly string GraphRootUri = ConfigurationManager.AppSettings["GraphRootUri"];

        public async Task ProcessCreateNewRequest(IDialogContext context, List<NewTeamDetails> teamDetailsList, string token)
        {
            foreach (var teamDetails in teamDetailsList)
            {
                var groupId = await CreateGroupAsyn(token, teamDetails.TeamName);
                if (IsValidGuid(groupId))
                {
                    await context.PostAsync($"Created O365 group for '{teamDetails.TeamName}'. Now, creating team which may take some time.");

                    var retryCount = 4;
                    string teamId = null;
                    while (retryCount > 0)
                    {
                        teamId = await CreateTeamAsyn(token, groupId);
                        if (IsValidGuid(teamId))
                        {
                            await context.PostAsync($" '{teamDetails.TeamName}' Team created successfully.");
                            break;
                        }
                        else
                        {
                            teamId = null;
                        }
                        retryCount--;
                        await Task.Delay(9000);
                    }

                    await CreateTeamAndChannels(context, token, teamDetails, teamId);
                }
                else
                {
                    await context.PostAsync($"Failed to create O365 Group due to internal error. Please try again later.");
                }
            }
        }

        public async Task ProcessUpdateRequest(IDialogContext context, List<NewTeamDetails> teamDetailsList, string token)
        {
            foreach (var teamDetails in teamDetailsList)
            {
                var teamId = await GetGroupId(token, teamDetails.TeamName);
                if (IsValidGuid(teamId))
                {
                    await CreateTeamAndChannels(context, token, teamDetails, teamId);
                }
                else
                {
                    await context.PostAsync($"Unable to find '{teamDetails.TeamName}' Team due to internal error. Please check team name try again later.");
                }
            }
        }

        private async Task CreateTeamAndChannels(IDialogContext context, string token, NewTeamDetails teamDetails, string teamId)
        {
            if (teamId != null)
            {
                foreach (var channelName in teamDetails.ChannelNames)
                {
                    var channelId = await CreateChannel(token, teamId, channelName, channelName);
                    if (String.IsNullOrEmpty(channelId))
                        await context.PostAsync($"Failed to create '{channelName}' channel in '{teamDetails.TeamName}' team.");
                }

                // Add users:
                foreach (var memberEmailId in teamDetails.MemberEmails)
                {
                    var result = await AddUserToTeam(token, teamId, memberEmailId);

                    if (!result)
                        await context.PostAsync($"Failed to add {memberEmailId} to {teamDetails.TeamName}. Check if user is already part of this team.");
                }

                await context.PostAsync($"Channels, Members Added successfully for '{teamDetails.TeamName}' team.");
            }
            else
            {
                await context.PostAsync($"Failed to create team due to internal error. Please try again later.");
            }
        }

        private static async Task ReplyWithMessage(Activity activity, ConnectorClient connector, string message)
        {
            var reply = activity.CreateReply();
            reply.Text = message;
            await connector.Conversations.ReplyToActivityAsync(reply);
        }

        private async Task<bool> AddUserToTeam(string token, string teamId, string userEmailId)
        {
            var userId = await GetUserId(token, userEmailId);
            return await AddTeamMemberAsync(token, teamId, userId);
        }

        bool IsValidGuid(string guid)
        {
            Guid teamGUID;
            return Guid.TryParse(guid, out teamGUID);
        }

        public async Task<string> CreateChannel(
            string accessToken, string teamId, string channelName, string channelDescription)
        {
            string endpoint = GraphRootUri + $"groups/{teamId}/team/channels";

            ChannelInfoBody channelInfo = new ChannelInfoBody()
            {
                description = channelDescription,
                displayName = channelName
            };

            return await PostRequest(accessToken, endpoint, JsonConvert.SerializeObject(channelInfo));
        }

        public async Task<string> CreateGroupAsyn(
            string accessToken, string groupName)
        {
            string endpoint = GraphRootUri + "groups/";

            GroupInfo groupInfo = new GroupInfo()
            {
                description = "Team for " + groupName,
                displayName = groupName,
                groupTypes = new string[] { "Unified" },
                mailEnabled = true,
                mailNickname = groupName.Replace(" ", "").Replace("-", "") + DateTime.Now.Second,
                securityEnabled = true
            };

            return await PostRequest(accessToken, endpoint, JsonConvert.SerializeObject(groupInfo));
        }


        public async Task<bool> AddTeamMemberAsync(
            string accessToken, string teamId, string userId)
        {
            string endpoint = GraphRootUri + $"groups/{teamId}/members/$ref";

            var userData = $"{{ \"@odata.id\": \"https://graph.microsoft.com/beta/directoryObjects/{userId}\" }}";

            using (var client = new HttpClient())
            {
                using (var request = new HttpRequestMessage(HttpMethod.Post, endpoint))
                {
                    request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                    request.Content = new StringContent(userData, Encoding.UTF8, "application/json");

                    using (HttpResponseMessage response = await client.SendAsync(request))
                    {
                        if (response.IsSuccessStatusCode)
                        {

                            return true;
                        }
                        return false;
                    }
                }
            }
        }


        public async Task<string> CreateTeamAsyn(
           string accessToken, string groupId)
        {
            // This might need Retries.
            string endpoint = GraphRootUri + $"groups/{groupId}/team";

            TeamSettings teamInfo = new TeamSettings()
            {
                funSettings = new Funsettings() { allowGiphy = true, giphyContentRating = "strict" },
                messagingSettings = new Messagingsettings() { allowUserEditMessages = true, allowUserDeleteMessages = true },
                memberSettings = new Membersettings() { allowCreateUpdateChannels = true }
            };
            return await PutRequest(accessToken, endpoint, JsonConvert.SerializeObject(teamInfo));
        }

        private static async Task<string> PostRequest(string accessToken, string endpoint, string groupInfo)
        {
            using (var client = new HttpClient())
            {
                using (var request = new HttpRequestMessage(HttpMethod.Post, endpoint))
                {
                    request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                    request.Content = new StringContent(groupInfo, Encoding.UTF8, "application/json");

                    using (HttpResponseMessage response = await client.SendAsync(request))
                    {
                        if (response.IsSuccessStatusCode)
                        {

                            var createdGroupInfo = JsonConvert.DeserializeObject<ResponseData>(response.Content.ReadAsStringAsync().Result);
                            return createdGroupInfo.id;
                        }
                        return null;
                    }
                }
            }
        }

        private static async Task<string> PutRequest(string accessToken, string endpoint, string groupInfo)
        {
            using (var client = new HttpClient())
            {
                using (var request = new HttpRequestMessage(HttpMethod.Put, endpoint))
                {
                    request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                    request.Content = new StringContent(groupInfo, Encoding.UTF8, "application/json");

                    using (HttpResponseMessage response = await client.SendAsync(request))
                    {
                        if (response.IsSuccessStatusCode)
                        {

                            var createdGroupInfo = JsonConvert.DeserializeObject<ResponseData>(response.Content.ReadAsStringAsync().Result);
                            return createdGroupInfo.id;
                        }
                        return null;
                    }
                }
            }
        }

        /// <summary>
        /// Get the current user's id from their profile.
        /// </summary>
        /// <param name="accessToken">Access token to validate user</param>
        /// <returns></returns>
        public async Task<string> GetGroupId(string accessToken, string teamName)
        {
            string endpoint = GraphRootUri + $"/groups?$filter=displayName eq '{teamName}'&$select=id";

            using (var client = new HttpClient())
            {
                using (var request = new HttpRequestMessage(HttpMethod.Get, endpoint))
                {
                    request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                    string groupId = "";
                    using (HttpResponseMessage response = await client.SendAsync(request))
                    {
                        if (response.IsSuccessStatusCode)
                        {
                            var json = JObject.Parse(await response.Content.ReadAsStringAsync());
                            try
                            {
                                groupId = json["value"].First["id"].ToString();
                            }
                            catch (Exception)
                            {
                                // Handle edge case.
                            }

                        }
                        return groupId?.Trim();
                    }
                }
            }
        }

        /// <summary>
        /// Get the current user's id from their profile.
        /// </summary>
        /// <param name="accessToken">Access token to validate user</param>
        /// <returns></returns>
        public async Task<string> GetUserId(string accessToken, string userEmailId)
        {
            string endpoint = GraphRootUri + $"users/{userEmailId}";
            string queryParameter = "?$select=id";

            using (var client = new HttpClient())
            {
                using (var request = new HttpRequestMessage(HttpMethod.Get, endpoint + queryParameter))
                {
                    request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                    string userId = "";
                    using (HttpResponseMessage response = await client.SendAsync(request))
                    {
                        if (response.IsSuccessStatusCode)
                        {
                            var json = JObject.Parse(await response.Content.ReadAsStringAsync());
                            userId = json.GetValue("id").ToString();
                        }
                        return userId?.Trim();
                    }
                }
            }
        }
    }
}