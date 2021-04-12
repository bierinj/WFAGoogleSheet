using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Http;
using Newtonsoft.Json; // Install Newtonsoft.Json with NuGet
using System.Net;
using System.IO;

namespace WFAGoolgeSheet
{
    public partial class Translator : Component
    {
        private static readonly string subscriptionKey = "0d6e23c01d284d5082a0a87aeb2de83e";
        private static readonly string endpoint = "https://api.cognitive.microsofttranslator.com/";
        // Add your location, also known as region. The default is global.
        // This is required if using a Cognitive Services resource.
        private static readonly string location = "brazilsouth";
        HttpWebRequestHandler HttpWebrequest = new HttpWebRequestHandler();
        public async Task<string> TranslationAsync(string lText)
        {

            if (!string.IsNullOrEmpty(lText))
            {
                string applicationid = "0d6e23c01d284d5082a0a87aeb2de83e";  // API key

            }
            return (null);
        }
    }
}

