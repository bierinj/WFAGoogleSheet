using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WFAGoolgeSheet
{
    public partial class GPSgeofence : Component
    {
        // static GPS fence values
        // and Bing Maps values
        public string bkey = "AhbjdGZqctwmlxK6GXWgkfE5CL7J2c5OWuTCk7WaAy-AhbjdGZqctwmlxK6GXWgkfE5CL7J2c5OWuTCk7WaAy-";
        static readonly float[] polyX = { -0.32541f,
            -0.30631F,
            -0.30640F,
            -0.33163F,
            -0.33103F,
            -0.36622F,
            -0.37137F,
            -0.35979F,
            -0.34614F,
            -0.32545F,
            -0.31189F};  // horizontal coordinates of corners
        static readonly float[] polyY = { -78.58218f,
            -78.56296F,
            -78.55945F,
            -78.55515F,
            -78.55043F,
            -78.55086F,
            -78.55919F,
            -78.57481F,
            -78.58253F,
            -78.58271F,
            -78.56975F};                        // vertical coordinates of corners
        readonly int polyCorners = polyX.Length;            // how many corners the polygon has
        //float x, y = 0;                                   // point to be tested

        public GPSgeofence()
        {
            InitializeComponent();
        }

        public GPSgeofence(IContainer container)
        {
            container.Add(this);

            InitializeComponent();
        }

        //  Globals which should be set before calling this function:
        //
        //  int    polyCorners  =  how many corners the polygon has
        //  float  polyX[]      =  horizontal coordinates of corners
        //  float  polyY[]      =  vertical coordinates of corners
        //  float  x, y         =  point to be tested
        //
        //  (Globals are used in this example for purposes of speed.  Change as
        //  desired.)
        //
        //  The function will return YES if the point x,y is inside the polygon, or
        //  NO if it is not.  If the point is exactly on the edge of the polygon,
        //  then the function may return YES or NO.
        //
        //  Note that division by zero is avoided because the division is protected
        //  by the "if" clause which surrounds it.

        public bool PointInPolygon(float x, float y)
        {
            int i, j = polyCorners - 1;
            bool oddNodes = false;

            for (i = 0; i < polyCorners; i++)
            {
                if ((polyY[i] < y && polyY[j] >= y
                || polyY[j] < y && polyY[i] >= y)
                && (polyX[i] <= x || polyX[j] <= x))
                {
                    oddNodes ^= (polyX[i] + (y - polyY[i]) / (polyY[j] - polyY[i]) * (polyX[j] - polyX[i]) < x);
                }
                j = i;
            }
            return oddNodes;
        }
    }
}
