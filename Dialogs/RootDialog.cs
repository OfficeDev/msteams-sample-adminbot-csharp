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
using Microsoft.Bot.Connector.Teams.Models;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using TeamsAdmin.Helper;

namespace Microsoft.Bot.Sample.TeamsAdmin.Dialogs
{
    [Serializable]
    public class RootDialog : IDialog<object>
    {
        private static readonly string ConnectionName = ConfigurationManager.AppSettings["ConnectionName"];
        private const string LastAction = "LastAction";
        private static readonly string LastCommand = string.Empty;

        public async Task StartAsync(IDialogContext context)
        {
            context.Wait(MessageReceivedAsync);
        }

        public async Task MessageReceivedAsync(IDialogContext context, IAwaitable<IMessageActivity> argument)
        {
            var activity = await argument as Activity;

            if (activity.Text == null)
                activity.Text = string.Empty;

            var message = Connector.Teams.ActivityExtensions.GetTextWithoutMentions(activity).ToLowerInvariant().Trim();

            if (message.Equals("help") || message.Equals("hi") || message.Equals("hello"))
            {
                await SendHelpMessage(context, activity);
            }
            else
            {
                // Check for file upload.
                if (activity.Attachments != null && activity.Attachments.Any())
                {
                    var token = await context.GetUserTokenAsync(ConnectionName).ConfigureAwait(false);
                    if (token == null || token.Token == null)
                    {
                        await SendOAuthCardAsync(context, activity);
                        return;
                    }
                    try
                    {
                        var attachment = activity.Attachments.First();
                        await HandleExcelAttachement(context, activity, token, attachment);
                    }
                    catch (Exception ex)
                    {
                        await context.PostAsync(ex.ToString());
                    }
                }
                else
                {
                    // All the commands can be executed...
                    switch (message)
                    {
                        case "create team":
                            context.UserData.SetValue(LastAction, message);
                            await CreateTeam(context, activity);
                            break;
                        case "add members":
                        case "add channels":
                        case "add members/channels":
                            context.UserData.SetValue(LastAction, message);
                            await UpdateTeam(context, activity);
                            break;
                        case "logout":
                            await Signout(context);
                            break;
                        default:
                            await context.PostAsync("Please check type help commands to know options.");
                            break;
                    }
                }

            }
        }

        #region Action Handlers

        private static async Task HandleExcelAttachement(IDialogContext context, Activity activity, TokenResponse token, Attachment attachment)
        {
            if (attachment.ContentType == FileDownloadInfo.ContentType)
            {
                FileDownloadInfo downloadInfo = (attachment.Content as JObject).ToObject<FileDownloadInfo>();
                var filePath = System.Web.Hosting.HostingEnvironment.MapPath("~/Files/");
                if (!Directory.Exists(filePath))
                    Directory.CreateDirectory(filePath);

                filePath += attachment.Name + DateTime.Now.Millisecond; // just to avoid name collision with other users. 
                if (downloadInfo != null)
                {
                    using (WebClient myWebClient = new WebClient())
                    {
                        // Download the Web resource and save it into the current filesystem folder.
                        myWebClient.DownloadFile(downloadInfo.DownloadUrl, filePath);

                    }
                    if (File.Exists(filePath))
                    {
                        var teamDetails = ExcelHelper.GetAddTeamDetails(filePath);
                        if (teamDetails == null)
                        {
                            await context.PostAsync($"Attachment received but unfortunately we are not able to read your excel file. Please make sure that all the colums are correct.");
                        }
                        else
                        {
                            string lastAction;
                            if (context.UserData.TryGetValue(LastAction, out lastAction))
                            {
                                await context.PostAsync($"Attachment received. Working on getting your {teamDetails.Count} Teams ready.");

                                GraphAPIHelper helper = new GraphAPIHelper();
                                if (lastAction == "create team")
                                {
                                    await helper.ProcessCreateNewRequest(context, teamDetails, token.Token);
                                }
                                else
                                {
                                    await helper.ProcessUpdateRequest(context, teamDetails, token.Token);
                                }
                            }
                            else
                            {
                                await context.PostAsync($"Not able to process your file. Please restart the flow.");
                            }
                            await SendHelpMessage(context, activity);
                        }

                        File.Delete(filePath);
                    }
                }
            }
        }

        private async Task CreateTeam(IDialogContext context, Activity activity)
        {
            var token = await context.GetUserTokenAsync(ConnectionName).ConfigureAwait(false);
            if (token != null && token.Token != null)
            {
                Activity reply = activity.CreateReply();

                ThumbnailCard card = GetThumbnailForTeamsAction();
                card.Title = "Create a new team";
                card.Subtitle = "Automate team creation by sharing team details";

                reply.TextFormat = TextFormatTypes.Xml;
                reply.Attachments.Add(card.ToAttachment());
                await context.PostAsync(reply);
            }
            else
            {
                await SendOAuthCardAsync(context, activity);
            }
        }

        private async Task UpdateTeam(IDialogContext context, Activity activity)
        {
            var token = await context.GetUserTokenAsync(ConnectionName).ConfigureAwait(false);
            if (token != null && token.Token != null)
            {
                Activity reply = activity.CreateReply();
                ThumbnailCard card = GetThumbnailForTeamsAction();

                card.Title = "Update existing team";
                card.Subtitle = "Automate adding members/channels by sharing team details";

                reply.TextFormat = TextFormatTypes.Xml;
                reply.Attachments.Add(card.ToAttachment());
                await context.PostAsync(reply);
            }
            else
            {
                await SendOAuthCardAsync(context, activity);
            }
        }
        #endregion

        #region Static Helpers

        private static ThumbnailCard GetThumbnailForTeamsAction()
        {
            return new ThumbnailCard
            {
                Text = @"Please go ahead and upload the excel file with team details in following format:  
                        <ol>
                        <li><strong>Team Name</strong>: String eg: <pre>IT Helpline</pre></li>
                        <li><strong>Channels</strong> : Comma separated channel names eg: <pre>my channel 1,my channel 2</pre></li>
                        <li><strong>Members</strong>  : Comma separated user emails eg: <pre>user1@org.com, user2@org.com</pre></li></ol>
                         </br> <strong>Note: Please keep first row header as described above. You can provide details for multiple teams row by row. Members/Channels columns can be empty.</strong>",
                Buttons = new List<CardAction>(),
            };
        }

        /// <summary>
        /// Signs the user out from AAD
        /// </summary>
        public static async Task Signout(IDialogContext context)
        {
            await context.SignOutUserAsync(ConnectionName);
            await context.PostAsync($"You have been signed out.");
        }

        public static async Task SendHelpMessage(IDialogContext context, Activity activity)
        {
            Activity reply = activity.CreateReply();
            ThumbnailCard card = GetHelpMessage();

            reply.TextFormat = TextFormatTypes.Xml;
            reply.Attachments.Add(card.ToAttachment());
            await context.PostAsync(reply);
        }

        internal static ThumbnailCard GetHelpMessage()
        {
            ThumbnailCard card = new ThumbnailCard
            {
                Title = "Welcome to Teams Creation Bot",
                Subtitle = "Your aide in creating & managing teams",
                Text = @"Use the bot for following 
                        <ol><li>Create a new team by uploading excel file with memeber details</li><li>Add new members to an existing team</li></ol>",
                Buttons = new List<CardAction>(),
            };

            card.Buttons.Add(new CardAction
            {
                Title = "Create a new team",
                DisplayText = "Create Team",
                Type = ActionTypes.MessageBack,
                Text = "Create Team",
                Value = "Create Team"

            });

            card.Buttons.Add(new CardAction
            {
                Title = "Add Members/Channels to existing team",
                DisplayText = "Add Members/Channels",
                Type = ActionTypes.MessageBack,
                Text = "Add Members",
                Value = "Add New Members"

            });
            return card;
        }
        #endregion

        #region Sign In Flow
        private async Task SendOAuthCardAsync(IDialogContext context, Activity activity)
        {
            var reply = await context.Activity.CreateOAuthReplyAsync(ConnectionName, "To do this, you'll first need to sign in.", "Sign In", true).ConfigureAwait(false);
            await context.PostAsync(reply);

            context.Wait(WaitForToken);
        }

        private async Task WaitForToken(IDialogContext context, IAwaitable<object> result)
        {
            var activity = await result as Activity;

            var tokenResponse = activity.ReadTokenResponseContent();
            if (tokenResponse != null)
            {
                // Use the token to do exciting things!

            }
            else
            {
                // Get the Activity Message as well as activity.value in case of Auto closing of pop-up
                string input = activity.Type == ActivityTypes.Message ? Connector.Teams.ActivityExtensions.GetTextWithoutMentions(activity)
                                                                : ((dynamic)(activity.Value)).state.ToString();
                if (!string.IsNullOrEmpty(input))
                {
                    tokenResponse = await context.GetUserTokenAsync(ConnectionName, input.Trim());
                    if (tokenResponse != null && tokenResponse.Token != null)
                    {
                        try
                        {
                            await context.PostAsync($"You are successfully signed in. Now, you can use create team command.");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex);
                        }
                    }
                    else
                    {
                        await context.PostAsync($"Hmm. Something went wrong. Let's try again.");
                    }
                    context.Wait(MessageReceivedAsync);
                    return;
                }
                await context.PostAsync($"Hmm. Something went wrong. Let's try again.");
                await SendOAuthCardAsync(context, activity);
            }
        }
        #endregion
    }
}