using GMap.NET.WindowsForms;
using GMap.NET.WindowsForms.Markers;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Documents;
//using System.Windows.Documents;
using System.Windows.Forms;

namespace WFAGoolgeSheet
{
    public partial class Form4: Form
    {
        protected override bool ShowWithoutActivation
        {
            get { return true; }
        }

        double lat;
        double lon;

        static GMapOverlay markersOverlay = null;
        static GMarkerGoogle marker = null;

        public Form4()
        {
            InitializeComponent();
            this.Activated += new EventHandler(Sub_LostFocus);
            this.Deactivate += new EventHandler(Sub_LostFocus);
        }

        private void Sub_GotFocus(object sender, EventArgs e)
        {
            //MessageBox.Show("'" + this.Text + "' got focus");
        }
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
        private void Form4_Load(object sender, EventArgs e)
        {
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

        private void label1_Click(object sender, EventArgs e)
        {

        }

        public void LoadIntoMap(double latitude, double longitude)
        {
            Form1 form1 = new Form1();
            //map.MapProvider = GMap.NET.MapProviders.GMapProviders.BingHybridMap;
            lat = latitude;
            lon = longitude;
            map.DragButton = MouseButtons.Left;
            map.MouseWheelZoomEnabled = true;

                // Initialize map:
                map.MapProvider = GMap.NET.MapProviders.BingMapProvider.Instance;
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

                GMapPolygon polygon = new GMapPolygon((List <GMap.NET.PointLatLng>)points, "mypolygon");
                polygon.Fill = new SolidBrush(Color.FromArgb(50, Color.Red));
                polygon.Stroke = new Pen(Color.Red, 1);
                polyOverlay.Polygons.Add(polygon);
                map.Overlays.Add(polyOverlay);
                markersOverlay = new GMapOverlay("markers");
                Program.formisup = true;
            map.MinZoom = 1;
            map.MaxZoom = 18;
            map.Zoom = 10;
          
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

        private void OnFormExit(object sender, FormClosingEventArgs e)
        {
            Form1 form1 = new Form1();
            Program.formisup = false;
            var zoomlevel = map.Zoom;
            Console.WriteLine(zoomlevel);
        }

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

            // don't forget to save the settings
            Properties.Settings.Default.Save();
        }
    }
}
