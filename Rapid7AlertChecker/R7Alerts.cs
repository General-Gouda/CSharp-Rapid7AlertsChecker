using System;
using HtmlAgilityPack;
using Html2Markdown;
using System.Collections.Generic;
using System.Timers;
using NLog;
using System.Configuration;

namespace Rapid7AlertChecker
{
    class R7AlertsService
    {
        private static Logger logger = LogManager.GetCurrentClassLogger();
        private static string graph_api_endpoint = ConfigurationManager.AppSettings["Graph_API_Endpoint"];
        private static string username = ConfigurationManager.AppSettings["Username"];
        private static Boolean testing = Convert.ToBoolean(ConfigurationManager.AppSettings["Testing"]);
        readonly Timer _timer;

        public R7AlertsService()
        {
            if (testing)
            {
                _timer = new Timer(Convert.ToUInt16(ConfigurationManager.AppSettings["TestingTimerInterval"]));
            }
            else
            {
                _timer = new Timer(Convert.ToUInt16(ConfigurationManager.AppSettings["TimerInterval"]));
            }

            _timer.Elapsed += new ElapsedEventHandler(R7Alerts);
        }

        public bool Start()
        {
            _timer.AutoReset = true;
            _timer.Enabled = true;
            _timer.Start();
            return true;
        }

        public bool Stop()
        {
            _timer.Stop();
            _timer.AutoReset = false;
            _timer.Enabled = false;
            return true;
        }

        public virtual void R7Alerts(object sender, ElapsedEventArgs e)
        {
            string slackAccess;

            if (testing)
            {
                slackAccess = "TestingSlackWebhookGoesHere"; // Testing
                username = ConfigurationManager.AppSettings["TestUsername"];
            }
            else
            {
                slackAccess = "ProductionSlackWebhookGoesHere"; // Production
            }

            SlackClient<string> slackClient = new SlackClient<string>(slackAccess);

            try
            {
                string dateTime = DateTime.Now.ToString();

                var adalResults = ServicePrincipal.GetS2SAccessTokenForProdMSAAsync().GetAwaiter().GetResult();

                string token = null;

                token = adalResults.AccessToken;

                if (token != null)
                {
                    dynamic mailFolders = null;

                    mailFolders = ApiCallers.Get_API_Results(
                        graph_api_endpoint,
                        string.Format("/users/{0}/mailFolders", username),
                        token
                    );

                    string inboxMailFolderID = null;
                    string archiveMailFolderID = null;

                    if (mailFolders != null)
                    {
                        foreach (var mailFolder in mailFolders)
                        {
                            if (mailFolder["displayName"] == "Inbox")
                            {
                                inboxMailFolderID = mailFolder["id"];
                            }
                            else if (mailFolder["displayName"] == "Archive")
                            {
                                archiveMailFolderID = mailFolder["id"];
                            }
                        }

                        dynamic emailJsonData = null;

                        emailJsonData = ApiCallers.Get_API_Results(
                                graph_api_endpoint,
                                string.Format("/users/{0}/mailFolders/{1}/messages?$top=200", username, inboxMailFolderID),
                                token
                            );

                        if (emailJsonData != null && emailJsonData.Count >= 1)
                        {
                            foreach (var email in emailJsonData)
                            {
                                string emailID = Convert.ToString(email["id"]);

                                if (email["isRead"] == false)
                                {
                                    EmailObject emailObject = new EmailObject()
                                    {
                                        BodyPreview = Convert.ToString(email["bodyPreview"]),
                                        EmailSubject = Convert.ToString(email["subject"]),
                                        EmailSentDateTime = Convert.ToDateTime(email["sentDateTime"]),
                                        EmailBodyHtml = Convert.ToString(email["body"]["content"]),
                                        EmailSentFrom = Convert.ToString(email["sender"]["emailAddress"]["address"])
                                    };

                                    HtmlDocument htmlDocument = new HtmlDocument();

                                    htmlDocument.LoadHtml(emailObject.EmailBodyHtml);

                                    HtmlNode htmlBody = htmlDocument.DocumentNode.SelectSingleNode("//body");

                                    Converter converter = new Converter();

                                    string htmlMarkdown;

                                    if (htmlBody != null)
                                    {
                                        try
                                        {
                                            htmlMarkdown = converter.Convert(htmlBody.InnerHtml);
                                        }
                                        catch (Exception)
                                        {
                                            htmlMarkdown = htmlBody.InnerText;
                                        }
                                    }
                                    else
                                    {
                                        htmlMarkdown = converter.Convert(emailObject.EmailBodyHtml);
                                    }

                                    string slackColor;
                                    string slackChannel = "";

                                    if (testing)
                                    {
                                        switch (emailObject.EmailSentFrom)
                                        {
                                            case "insightphish@rapid7.com":
                                                slackColor = SlackColors.Amber;
                                                break;
                                            case "insight_noreply@rapid7.com":
                                                slackColor = SlackColors.Red;
                                                break;
                                            default:
                                                slackColor = SlackColors.White;
                                                break;
                                        }
                                    }
                                    else
                                    {
                                        switch (emailObject.EmailSentFrom)
                                        {
                                            case "insightphish@rapid7.com":
                                                slackColor = SlackColors.Amber;
                                                slackChannel = "r7_phish_alerts"; // This can be changed to another channel name
                                                break;
                                            case "insight_noreply@rapid7.com":
                                                slackColor = SlackColors.Red;
                                                slackChannel = "r7_idr_alerts"; // This can be changed to another channel name
                                                break;
                                            default:
                                                slackColor = SlackColors.White;
                                                break;
                                        }
                                    }

                                    List<Attachment> slackAttachments = new List<Attachment>();

                                    Attachment attachment = new Attachment()
                                    {
                                        text = htmlMarkdown.Replace("**", "*"),
                                        color = slackColor,
                                        pretext = string.Format($"`{emailObject.EmailSubject}` sent on *{emailObject.EmailSentDateTime.ToLocalTime().ToShortDateString()}* at *{emailObject.EmailSentDateTime.ToLocalTime().TimeOfDay} EST*")
                                    };

                                    slackAttachments.Add(attachment);

                                    bool postMessage = slackClient.PostMessage(
                                        channel: slackChannel,
                                        username: "Rapid7 Alerts",
                                        attachments: slackAttachments
                                    );

                                    if (postMessage)
                                    {
                                        // Marks email as read
                                        ApiCallers.MarkAsRead(graph_api_endpoint, username, emailID, token);

                                        // Moves email to Archive folder in mailbox
                                        ApiCallers.MoveEmail(archiveMailFolderID, graph_api_endpoint, username, emailID, token);
                                    }
                                    else
                                    {
                                        logger.Info("Failed to post message. Leaving it unread and in the Inbox.");
                                    }
                                }
                                else
                                {
                                    // Moves email to Archive folder in mailbox
                                    ApiCallers.MoveEmail(archiveMailFolderID, graph_api_endpoint, username, emailID, token);
                                }
                            }
                        }
                    }
                }
            }
            catch (System.NullReferenceException nullRef)
            {
                logger.Error($"Null Reference Exception: {nullRef}");
            }
            catch (Exception ex)
            {
                List<Attachment> slackErrorAttachments = new List<Attachment>();

                Attachment errorAttachment = new Attachment()
                {
                    text = ex.ToString(),
                    color = SlackColors.Purple,
                    pretext = "*Program Exception - Halting service*"
                };

                slackErrorAttachments.Add(errorAttachment);

                slackClient.PostMessage(
                    username: "Rapid7 Alerts",
                    attachments: slackErrorAttachments
                );

                Stop();
            }
        }
    }
}
