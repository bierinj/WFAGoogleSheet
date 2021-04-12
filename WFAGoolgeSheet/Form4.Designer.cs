namespace WFAGoolgeSheet
{
    partial class Form4
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form4));
            this.map = new GMap.NET.WindowsForms.GMapControl();
            this.savZoom = new System.Windows.Forms.Button();
            this.savPos = new System.Windows.Forms.Button();
            this.zoomlvl = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // map
            // 
            resources.ApplyResources(this.map, "map");
            this.map.Bearing = 0F;
            this.map.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.map.CanDragMap = true;
            this.map.EmptyTileColor = System.Drawing.Color.Navy;
            this.map.GrayScaleMode = false;
            this.map.HelperLineOption = GMap.NET.WindowsForms.HelperLineOptions.DontShow;
            this.map.LevelsKeepInMemory = 5;
            this.map.MarkersEnabled = true;
            this.map.MaxZoom = 18;
            this.map.MinZoom = 2;
            this.map.MouseWheelZoomEnabled = true;
            this.map.MouseWheelZoomType = GMap.NET.MouseWheelZoomType.ViewCenter;
            this.map.Name = "map";
            this.map.NegativeMode = false;
            this.map.PolygonsEnabled = true;
            this.map.RetryLoadTile = 0;
            this.map.RoutesEnabled = true;
            this.map.ScaleMode = GMap.NET.WindowsForms.ScaleModes.Fractional;
            this.map.SelectedAreaFillColor = System.Drawing.Color.FromArgb(((int)(((byte)(33)))), ((int)(((byte)(65)))), ((int)(((byte)(105)))), ((int)(((byte)(225)))));
            this.map.ShowTileGridLines = false;
            this.map.Zoom = 5D;
            this.map.OnMapZoomChanged += new GMap.NET.MapZoomChanged(this.ZoomChange);
            // 
            // savZoom
            // 
            resources.ApplyResources(this.savZoom, "savZoom");
            this.savZoom.Name = "savZoom";
            this.savZoom.UseVisualStyleBackColor = true;
            this.savZoom.Click += new System.EventHandler(this.savZoom_Click);
            // 
            // savPos
            // 
            resources.ApplyResources(this.savPos, "savPos");
            this.savPos.Name = "savPos";
            this.savPos.UseVisualStyleBackColor = true;
            this.savPos.Click += new System.EventHandler(this.savPos_Click);
            // 
            // zoomlvl
            // 
            resources.ApplyResources(this.zoomlvl, "zoomlvl");
            this.zoomlvl.Name = "zoomlvl";
            this.zoomlvl.TextChanged += new System.EventHandler(this.zoomlvl_TextChanged);
            // 
            // label1
            // 
            resources.ApplyResources(this.label1, "label1");
            this.label1.Name = "label1";
            // 
            // comboBox1
            // 
            resources.ApplyResources(this.comboBox1, "comboBox1");
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Items.AddRange(new object[] {
            resources.GetString("comboBox1.Items"),
            resources.GetString("comboBox1.Items1"),
            resources.GetString("comboBox1.Items2"),
            resources.GetString("comboBox1.Items3"),
            resources.GetString("comboBox1.Items4"),
            resources.GetString("comboBox1.Items5")});
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.SelectedValueChanged += new System.EventHandler(this.comboBox1_SelectedValueChanged);
            // 
            // label2
            // 
            resources.ApplyResources(this.label2, "label2");
            this.label2.Name = "label2";
            // 
            // Form4
            // 
            resources.ApplyResources(this, "$this");
            this.AccessibleRole = System.Windows.Forms.AccessibleRole.ScrollBar;
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoValidate = System.Windows.Forms.AutoValidate.EnablePreventFocusChange;
            this.Controls.Add(this.label2);
            this.Controls.Add(this.comboBox1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.zoomlvl);
            this.Controls.Add(this.savPos);
            this.Controls.Add(this.savZoom);
            this.Controls.Add(this.map);
            this.Name = "Form4";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show;
            this.TopMost = true;
            this.Deactivate += new System.EventHandler(this.Sub_LostFocus);
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.OnFormExit);
            this.Load += new System.EventHandler(this.Form4_Load);
            this.Leave += new System.EventHandler(this.Sub_LostFocus);
            this.Resize += new System.EventHandler(this.formSizeChange);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        public GMap.NET.WindowsForms.GMapControl map;
        private System.Windows.Forms.Button savZoom;
        private System.Windows.Forms.Button savPos;
        private System.Windows.Forms.TextBox zoomlvl;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.Label label2;
    }
}