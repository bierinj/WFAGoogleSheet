namespace WFAGoolgeSheet
{
    /// <summary>
    /// Defines the <see cref="GPSgeofenceBase" />.
    /// </summary>
    public abstract class GPSgeofenceBase
    {
        /// <summary>
        /// The pointInPolygon.
        /// </summary>
        /// <param name="x">The x<see cref="float"/>.</param>
        /// <param name="y">The y<see cref="float"/>.</param>
        /// <returns>The <see cref="bool"/>.</returns>
        public abstract bool pointInPolygon(float x, float y);
    }
}
