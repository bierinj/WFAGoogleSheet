namespace WFAGoolgeSheet
{
    using System;
    using System.Data.Services;
    using System.IO;
    using System.Net;
    using System.ServiceModel.Channels;
    using System.Threading;

    /// <summary>
    /// Defines the <see cref="HttpWebRequestHandler" />.
    /// </summary>
    public class HttpWebRequestHandler : IRequestHandler
    {
        /// <summary>
        /// Gets the RequestConstants.
        /// </summary>
        public object RequestConstants { get; private set; }

        /// <summary>
        /// The GetReply.
        /// </summary>
        /// <param name="url">The url<see cref="string"/>.</param>
        /// <returns>The <see cref="string"/>.</returns>
        public string GetReply(string url)
        {
            var content = string.Empty;
            var request = (HttpWebRequest)WebRequest.Create(url);
            request.UserAgent = ".NET Framework Test Client";
            request.Method = "GET";

            for (int retryMax = 3; retryMax > 0; retryMax--)
            {
                try
                {
                    using (var response = (HttpWebResponse)request.GetResponse())
                    {
                        using (var stream = response.GetResponseStream())
                        {
                            using (var sr = new StreamReader(stream))
                            {
                                content = sr.ReadToEnd();
                            }
                        }
                    }
                    return content;
                }


                catch (WebException e)
                {
                    if (e.Status == WebExceptionStatus.Timeout)
                    {
                        Thread.Sleep(50);
                    }
                    //else throw;
                }
            }

            return content;
        }

        /// <summary>
        /// The ProcessRequestForMessage.
        /// </summary>
        /// <param name="messageBody">The messageBody<see cref="Stream"/>.</param>
        /// <returns>The <see cref="Message"/>.</returns>
        public Message ProcessRequestForMessage(Stream messageBody)
        {
            throw new NotImplementedException();
        }
    }
}
