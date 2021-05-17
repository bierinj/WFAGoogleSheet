namespace WFAGoolgeSheet
{
    using System.ComponentModel;

    /// <summary>
    /// Defines the <see cref="GPSgeocode" />.
    /// </summary>
    public partial class GPSgeocode : Component
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="GPSgeocode"/> class.
        /// </summary>
        public GPSgeocode()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="GPSgeocode"/> class.
        /// </summary>
        /// <param name="container">The container<see cref="IContainer"/>.</param>
        public GPSgeocode(IContainer container)
        {
            container.Add(this);

            InitializeComponent();
        }
    }
}
