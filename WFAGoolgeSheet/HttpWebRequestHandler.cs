using System;
using System.ServiceModel;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Data.Services;
using System.Net;
using System.IO;
using System.ServiceModel.Channels;

namespace WFAGoolgeSheet
{
        public class HttpWebRequestHandler : IRequestHandler
        {
        public object RequestConstants { get; private set; }

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
                    else throw;
                }
            }
            
            return content;
        }
        public Message ProcessRequestForMessage(Stream messageBody)
        {
            throw new NotImplementedException();
        }
    }
}

