namespace WFAGoolgeSheet
{
    using System.ComponentModel;
    using System.Threading.Tasks;

    /// <summary>
    /// Defines the <see cref="Translator" />.
    /// </summary>
    public partial class Translator : Component
    {
        /// <summary>
        /// Defines the subscriptionKey.
        /// </summary>
        private static readonly string subscriptionKey = " ";

        /// <summary>
        /// Defines the endpoint.
        /// </summary>
        private static readonly string endpoint = "https://api.cognitive.microsofttranslator.com/";

        // Add your location, also known as region. The default is global.
        // This is required if using a Cognitive Services resource.
        /// <summary>
        /// Defines the location.
        /// </summary>
        private static readonly string location = "brazilsouth";

        /// <summary>
        /// Defines the HttpWebrequest.
        /// </summary>
        internal HttpWebRequestHandler HttpWebrequest = new HttpWebRequestHandler();

        /// <summary>
        /// The TranslationAsync.
        /// </summary>
        /// <param name="lText">The lText<see cref="string"/>.</param>
        /// <returns>The <see cref="Task{string}"/>.</returns>
        public async Task<string> TranslationAsync(string lText)
        {

            if (!string.IsNullOrEmpty(lText))
            {
                string applicationid = " ";  // API key
            }
            return (null);
        }
    }
}
