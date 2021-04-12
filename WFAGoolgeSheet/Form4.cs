namespace WFAGoolgeSheet
{
    using GMap.NET.WindowsForms;
    using GMap.NET.WindowsForms.Markers;
    using Microsoft.Win32;
    using System;
    using System.Collections.Generic;
    using System.Drawing;
    using System.Windows.Forms;

    /// <summary>
    /// Defines the <see cref="Form4" />.
    /// </summary>
    public partial class Form4 : Form
    {
        /// <summary>
        /// Gets a value indicating whether ShowWithoutActivation.
        /// </summary>
        protected override bool ShowWithoutActivation
        {
            get { return true; }
        }

        /// <summary>
        /// Gets the mapSource.
        /// </summary>
        public string mapSource { get; private set; }

        /// <summary>
        /// Defines the lat.
        /// </summary>
        internal double lat;

        /// <summary>
        /// Defines the lon.
        /// </summary>
        internal double lon;

        /// <summary>
        /// Defines the markersOverlay.
        /// </summary>
        internal static GMapOverlay markersOverlay = null;

        /// <summary>
        /// Defines the marker.
        /// </summary>
        internal static GMarkerGoogle marker = null;

        /// <summary>
        /// Defines the textupdate.
        /// </summary>
        internal bool textupdate = false;

        /// <summary>
        /// Initializes a new instance of the <see cref="Form4"/> class.
        /// </summary>
        public Form4()
        {
            InitializeComponent();

            //this.Font = System.Drawing.SystemFonts.IconTitleFont;
            //SystemEvents.UserPreferenceChanged += new UserPreferenceChangedEventHandler(SystemEvents_UserPreferenceChanged);
            //this.FormClosing += new FormClosingEventHandler(Form4_FormClosing);

            this.Activated += new EventHandler(Sub_LostFocus);
            this.Deactivate += new EventHandler(Sub_LostFocus);
        }

        /// <summary>
        /// The Sub_GotFocus.
        /// </summary>
        /// <param name="sender">The sender<see cref="object"/>.</param>
        /// <param name="e">The e<see cref="EventArgs"/>.</param>
        private void Sub_GotFocus(object sender, EventArgs e)
        {
        }

        /// <summary>
        /// The Sub_LostFocus.
        /// </summary>
        /// <param name="sender">The sender<see cref="object"/>.</param>
        /// <param name="e">The e<see cref="EventArgs"/>.</param>
        private void Sub_LostFocus(object sender, EventArgs e)
        {
            Form4 form4 = new Form4();
            //MessageBox.Show("'" + this.Text + "' lost focus");

            //this.Show();
            form4.WindowState = FormWindowState.Normal;
            //form4.BringToFront();
            //form4.TopLevel = true;
            form4.Focus();
        }

        /// <summary>
        /// The SystemEvents_UserPreferenceChanged.
        /// </summary>
        /// <param name="sender">The sender<see cref="object"/>.</param>
        /// <param name="e">The e<see cref="UserPreferenceChangedEventArgs"/>.</param>
        internal void SystemEvents_UserPreferenceChanged(object sender, UserPreferenceChangedEventArgs e)
        {
            if (e.Category == UserPreferenceCategory.Window)
            {
                this.Font = System.Drawing.SystemFonts.IconTitleFont;
            }
        }

        /// <summary>
        /// The Form4_FormClosing.
        /// </summary>
        /// <param name="sender">The sender<see cref="object"/>.</param>
        /// <param name="e">The e<see cref="FormClosingEventArgs"/>.</param>
        internal void Form4_FormClosing(object sender, FormClosingEventArgs e)
        {
            SystemEvents.UserPreferenceChanged -= new UserPreferenceChangedEventHandler(SystemEvents_UserPreferenceChanged);
        }

        /// <summary>
        /// The Form4_Load.
        /// </summary>
        /// <param name="sender">The sender<see cref="object"/>.</param>
        /// <param name="e">The e<see cref="EventArgs"/>.</param>
        private void Form4_Load(object sender, EventArgs e)
        {
            Form1 form1 = new Form1();
            if (form1.radioButton5.Checked)
            {
                int i = form1.comboBox1.SelectedIndex;
                //ChangeLanguage to Ecuadorian Spanish
                string language = "es-EC";

                // Save user choice in settings
                Properties.Settings.Default.Language = "es-EC";
                Properties.Settings.Default.Save();

                form1.ChangeLanguage(language);
                form1.comboBox1.SelectedIndex = i;
                form1.Refresh();
            }
            if (form1.radioButton6.Checked)
            {
                int i = form1.comboBox1.SelectedIndex;
                //ChangeLanguage to American English
                string language = "en-US";

                // Save user choice in settings
                Properties.Settings.Default.Language = "en-US";
                Properties.Settings.Default.Save();

                form1.ChangeLanguage(language);
                form1.comboBox1.SelectedIndex = i;
                form1.Refresh();
            }
            if (Properties.Settings.Default.F1Size.Width == 0 || Properties.Settings.Default.F1Size.Height == 0)
            {
                // first start
                // optional: add default values
            }
            else
            {
                this.WindowState = Properties.Settings.Default.F1State;

                // we don't want a minimized window at startup
                if (this.WindowState == FormWindowState.Minimized) this.WindowState = FormWindowState.Normal;

                this.Location = Properties.Settings.Default.F1Location;
                this.Size = Properties.Settings.Default.F1Size;
            }
        }

        /// <summary>
        /// The label1_Click.
        /// </summary>
        /// <param name="sender">The sender<see cref="object"/>.</param>
        /// <param name="e">The e<see cref="EventArgs"/>.</param>
        private void label1_Click(object sender, EventArgs e)
        {
        }

        /// <summary>
        /// The LoadIntoMap.
        /// </summary>
        /// <param name="latitude">The latitude<see cref="double"/>.</param>
        /// <param name="longitude">The longitude<see cref="double"/>.</param>
        public void LoadIntoMap(double latitude, double longitude)
        {
            Form1 form1 = new Form1();
            //map.MapProvider = GMap.NET.MapProviders.GMapProviders.BingHybridMap;
            lat = latitude;
            lon = longitude;
            map.DragButton = MouseButtons.Left;
            map.MouseWheelZoomEnabled = true;

            // Initialize map:
            comboBox1.SelectedItem = Properties.Settings.Default.gMapSource;
            mapSource = (string)comboBox1.SelectedItem;
            switch (mapSource)
            {
                case "BingMaps":
                    map.MapProvider = GMap.NET.MapProviders.BingMapProvider.Instance;
                    break;
                case "GoogleMaps":
                    map.MapProvider = GMap.NET.MapProviders.GoogleMapProvider.Instance;
                    break;
                case "ArcGISMaps":
                    map.MapProvider = GMap.NET.MapProviders.ArcGIS_StreetMap_World_2D_MapProvider.Instance;
                    break;
                case "OpenStreetMap":
                    map.MapProvider = GMap.NET.MapProviders.OpenStreetMapProvider.Instance;
                    break;
                case "Yandex":
                    map.MapProvider = GMap.NET.MapProviders.YandexHybridMapProvider.Instance;
                    break;
                case "OpenCycleMaps":
                    map.MapProvider = GMap.NET.MapProviders.OpenCycleMapProvider.Instance;
                    break;
                default:
                    map.MapProvider = GMap.NET.MapProviders.BingMapProvider.Instance;
                    break;
            }

            //map.MapProvider = GMap.NET.MapProviders.OpenStreetMapQuestHybridProvider.Instance;
            GMap.NET.GMaps.Instance.Mode = GMap.NET.AccessMode.ServerOnly;
            object p = map.SetPositionByKeywords("Quito, Ecuador");

            // add GeoFence
            GPSgeofence gPSgeofence = new GPSgeofence();
            GPSgeofence fence = gPSgeofence;
            fence.ReadGPSfence();
            GMapOverlay polyOverlay = new GMapOverlay("polygons");
            IList<GMap.NET.PointLatLng> points = new List<GMap.NET.PointLatLng>();
            for (int i = 0; i < fence.polyCorners; i++)
            {
                double.TryParse(fence.polyX[i].ToString(), out lat);
                double.TryParse(fence.polyY[i].ToString(), out lon);
                points.Add(new GMap.NET.PointLatLng(lat, lon));
            }

            GMapPolygon polygon = new GMapPolygon((List<GMap.NET.PointLatLng>)points, "mypolygon");
            polygon.Fill = new SolidBrush(Color.FromArgb(50, Color.Red));
            polygon.Stroke = new Pen(Color.Red, 1);
            polyOverlay.Polygons.Add(polygon);
            map.Overlays.Add(polyOverlay);
            markersOverlay = new GMapOverlay("markers");
            Program.formisup = true;
            map.MinZoom = 1;
            map.MaxZoom = 18;

            map.Zoom = Properties.Settings.Default.gMapZoom;
            zoomlvl.Text = map.Zoom.ToString("0.00");
            lat = latitude;
            lon = longitude;
            //
            // add Marker
            //
            marker = new GMarkerGoogle(new GMap.NET.PointLatLng(lat, lon),
                  GMarkerGoogleType.green);
            map.Position = new GMap.NET.PointLatLng(lat, lon);
            markersOverlay.Markers.Add(marker);
            map.Overlays.Add(markersOverlay);
            map.ZoomAndCenterMarkers(markersOverlay.Id);
            map.Update();
        }

        /// <summary>
        /// The OnFormExit.
        /// </summary>
        /// <param name="sender">The sender<see cref="object"/>.</param>
        /// <param name="e">The e<see cref="FormClosingEventArgs"/>.</param>
        private void OnFormExit(object sender, FormClosingEventArgs e)
        {
            Form1 form1 = new Form1();
            Program.formisup = false;
        }

        /// <summary>
        /// The formSizeChange.
        /// </summary>
        /// <param name="sender">The sender<see cref="object"/>.</param>
        /// <param name="e">The e<see cref="EventArgs"/>.</param>
        private void formSizeChange(object sender, EventArgs e)
        {
            Properties.Settings.Default.F1State = this.WindowState;
            if (this.WindowState == FormWindowState.Normal)
            {
                // save location and size if the state is normal
                Properties.Settings.Default.F1Location = this.Location;
                Properties.Settings.Default.F1Size = this.Size;
            }
            else
            {
                // save the RestoreBounds if the form is minimized or maximized!
                Properties.Settings.Default.F1Location = this.RestoreBounds.Location;
                Properties.Settings.Default.F1Size = this.RestoreBounds.Size;
            }
        }

        /// <summary>
        /// The ZoomChange.
        /// </summary>
        private void ZoomChange()
        {
            textupdate = true;
            zoomlvl.Text = map.Zoom.ToString("0.00");
            //if (!textupdate)zoomlvl.Update();
            textupdate = false;
        }

        /// <summary>
        /// The savPos_Click.
        /// </summary>
        /// <param name="sender">The sender<see cref="object"/>.</param>
        /// <param name="e">The e<see cref="EventArgs"/>.</param>
        private void savPos_Click(object sender, EventArgs e)
        {
            formSizeChange(sender, e);
            // don't forget to save the settings
            Properties.Settings.Default.Save();
        }

        /// <summary>
        /// The savZoom_Click.
        /// </summary>
        /// <param name="sender">The sender<see cref="object"/>.</param>
        /// <param name="e">The e<see cref="EventArgs"/>.</param>
        private void savZoom_Click(object sender, EventArgs e)
        {
            //Properties.Settings.Default.gMapZoom = (float)map.Zoom;
            // don't forget to save the settings
            //Properties.Settings.Default.Save();
            this.Close();
        }

        /// <summary>
        /// The zoomlvl_TextChanged.
        /// </summary>
        /// <param name="sender">The sender<see cref="object"/>.</param>
        /// <param name="e">The e<see cref="EventArgs"/>.</param>
        private void zoomlvl_TextChanged(object sender, EventArgs e)
        {
            if (!textupdate)
            {
                map.ZoomAndCenterMarkers(markersOverlay.Id);
                map.Update();
            }
        }

        /// <summary>
        /// The comboBox1_SelectedValueChanged.
        /// </summary>
        /// <param name="sender">The sender<see cref="object"/>.</param>
        /// <param name="e">The e<see cref="EventArgs"/>.</param>
        private void comboBox1_SelectedValueChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.gMapSource = (string)comboBox1.SelectedItem;
            map.Refresh();
            map.Update();
            Properties.Settings.Default.Save();
            LoadIntoMap(lat, lon);
        }
    }
}
