using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Drawing;
using System.Runtime.InteropServices;
using System.IO;

namespace MTools2013
{
    class CoordPrompt: Form
    {
        public delegate void okButtonAction(string m1,string m2,string m3);
        public event okButtonAction okButtonActionEvent;
        
        Excel.Worksheet workSheet;
        Excel.Worksheet activWorkSheet = Globals.ThisAddIn.GetActiveWorksheet();
        int textBoxSelector;


        public CoordPrompt()
        {
            InitializeComponent();
        }


        public const int WM_NCLBUTTONDOWN = 0xA1;
        public const int HT_CAPTION = 0x2;
        [DllImportAttribute("user32.dll")]
        public static extern int SendMessage(System.IntPtr hWnd, int Msg, int wParam, int lParam);
        [DllImportAttribute("user32.dll")]
        public static extern bool ReleaseCapture();



        private void panel1_MouseDown(object sender, MouseEventArgs e)
        {
            textBoxSelector = 0;
            ReleaseCapture();
            SendMessage(this.Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
        }


        private void InitializeComponent()
        {
            this.titleLabel = new System.Windows.Forms.Label();
            this.xLabel = new System.Windows.Forms.Label();
            this.yLabel = new System.Windows.Forms.Label();
            this.distLabel = new System.Windows.Forms.Label();
            this.XcoordinateAdrsBox = new System.Windows.Forms.TextBox();
            this.YcoordinateAdrsBox = new System.Windows.Forms.TextBox();
            this.distAdrsBox = new System.Windows.Forms.TextBox();
            this.okButton = new System.Windows.Forms.Button();
            this.cancelButton = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.closeButton = new System.Windows.Forms.Button();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // titleLabel
            // 
            this.titleLabel.AutoSize = true;
            this.titleLabel.Location = new System.Drawing.Point(125, 10);
            this.titleLabel.Name = "titleLabel";
            this.titleLabel.Size = new System.Drawing.Size(152, 26);
            this.titleLabel.TabIndex = 0;
            this.titleLabel.Text = "COORDINATE GENERATOR \n (with formula)";
            this.titleLabel.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // xLabel
            // 
            this.xLabel.AutoSize = true;
            this.xLabel.Location = new System.Drawing.Point(24, 85);
            this.xLabel.Name = "xLabel";
            this.xLabel.Size = new System.Drawing.Size(48, 13);
            this.xLabel.TabIndex = 0;
            this.xLabel.Text = "X-Coord:";
            // 
            // yLabel
            // 
            this.yLabel.AutoSize = true;
            this.yLabel.Location = new System.Drawing.Point(144, 85);
            this.yLabel.Name = "yLabel";
            this.yLabel.Size = new System.Drawing.Size(48, 13);
            this.yLabel.TabIndex = 0;
            this.yLabel.Text = "Y-Coord:";
            // 
            // distLabel
            // 
            this.distLabel.AutoSize = true;
            this.distLabel.Location = new System.Drawing.Point(257, 85);
            this.distLabel.Name = "distLabel";
            this.distLabel.Size = new System.Drawing.Size(20, 13);
            this.distLabel.TabIndex = 0;
            this.distLabel.Text = "Distance:";
            // 
            // XcoordinateAdrsBox
            // 
            this.XcoordinateAdrsBox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.XcoordinateAdrsBox.Location = new System.Drawing.Point(76, 82);
            this.XcoordinateAdrsBox.Name = "XcoordinateAdrsBox";
            this.XcoordinateAdrsBox.Size = new System.Drawing.Size(60, 20);
            this.XcoordinateAdrsBox.TabIndex = 1;
            this.XcoordinateAdrsBox.GotFocus += new System.EventHandler(this.XcoordinateAdrsBox_GotFocus);
            // 
            // YcoordinateAdrsBox
            // 
            this.YcoordinateAdrsBox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.YcoordinateAdrsBox.Location = new System.Drawing.Point(195, 82);
            this.YcoordinateAdrsBox.Name = "YcoordinateAdrsBox";
            this.YcoordinateAdrsBox.Size = new System.Drawing.Size(60, 20);
            this.YcoordinateAdrsBox.TabIndex = 1;
            this.YcoordinateAdrsBox.GotFocus += new System.EventHandler(this.YcoordinateAdrsBox_GotFocus);
            // 
            // distAdrsBox
            // 
            this.distAdrsBox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.distAdrsBox.Location = new System.Drawing.Point(312, 82);
            this.distAdrsBox.Margin = new System.Windows.Forms.Padding(5, 3, 0, 0);
            this.distAdrsBox.Name = "distAdrsBox";
            this.distAdrsBox.Size = new System.Drawing.Size(60, 20);
            this.distAdrsBox.TabIndex = 1;
            this.distAdrsBox.GotFocus += new System.EventHandler(this.distAdrsBox_GotFocus);
            // 
            // okButton
            // 
            this.okButton.BackColor = System.Drawing.Color.DeepSkyBlue;
            this.okButton.FlatAppearance.BorderSize = 0;
            this.okButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.okButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.okButton.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.okButton.Location = new System.Drawing.Point(55, 140);
            this.okButton.Name = "okButton";
            this.okButton.Size = new System.Drawing.Size(80, 25);
            this.okButton.TabIndex = 2;
            this.okButton.Text = "OK";
            this.okButton.UseVisualStyleBackColor = false;
            this.okButton.Click += new System.EventHandler(this.okButton_Click);
            // 
            // cancelButton
            // 
            this.cancelButton.BackColor = System.Drawing.Color.DeepSkyBlue;
            this.cancelButton.FlatAppearance.BorderSize = 0;
            this.cancelButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.cancelButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cancelButton.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.cancelButton.Location = new System.Drawing.Point(265, 140);
            this.cancelButton.Name = "cancelButton";
            this.cancelButton.Size = new System.Drawing.Size(80, 25);
            this.cancelButton.TabIndex = 2;
            this.cancelButton.Text = "Cancel";
            this.cancelButton.UseVisualStyleBackColor = false;
            this.cancelButton.Click += new System.EventHandler(this.cancelButton_Click);
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.titleLabel);
            this.panel1.Controls.Add(this.xLabel);
            this.panel1.Controls.Add(this.yLabel);
            this.panel1.Controls.Add(this.distLabel);
            this.panel1.Controls.Add(this.XcoordinateAdrsBox);
            this.panel1.Controls.Add(this.YcoordinateAdrsBox);
            this.panel1.Controls.Add(this.distAdrsBox);
            this.panel1.Controls.Add(this.okButton);
            this.panel1.Controls.Add(this.cancelButton);
            this.panel1.Controls.Add(this.closeButton);
            this.panel1.Location = new System.Drawing.Point(1, 1);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(399, 199);
            this.panel1.TabIndex = 3;
            this.panel1.MouseDown += new System.Windows.Forms.MouseEventHandler(this.panel1_MouseDown);
            // 
            // closeButton
            // 
            this.closeButton.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(247)))), ((int)(((byte)(79)))), ((int)(((byte)(83)))));
            this.closeButton.FlatAppearance.BorderSize = 0;
            this.closeButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.closeButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.closeButton.Location = new System.Drawing.Point(368, -1);
            this.closeButton.Margin = new System.Windows.Forms.Padding(0);
            this.closeButton.Name = "closeButton";
            this.closeButton.Size = new System.Drawing.Size(30, 19);
            this.closeButton.TabIndex = 3;
            this.closeButton.Text = "X";
            this.closeButton.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            this.closeButton.UseVisualStyleBackColor = false;
            this.closeButton.Click += new System.EventHandler(this.closeButton_Click);
            // 
            // CoordPrompt
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.ClientSize = new System.Drawing.Size(400, 200);
            this.Controls.Add(this.panel1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "CoordPrompt";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Input Parameters";
            this.TopMost = true;
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);

        }
        private Label titleLabel; 
        private Label xLabel;
        private Label yLabel;
        private Label distLabel;
        private TextBox XcoordinateAdrsBox;
        private TextBox YcoordinateAdrsBox;
        private TextBox distAdrsBox;
        private Button okButton;
        private Button cancelButton;
        private Panel panel1;
        private Button closeButton;
        
        
        private void closeButton_Click(object sender, EventArgs e)
        {
            this.Dispose(true);
        }
        private void cancelButton_Click(object sender, EventArgs e)
        {
            this.Dispose(true);
        }

        private void okButton_Click(object sender, EventArgs e)
        {
            textBoxSelector = 0;

            if (XcoordinateAdrsBox.Text != "" & YcoordinateAdrsBox.Text != "" & distAdrsBox.Text != "")
            {
               string xRange = XcoordinateAdrsBox.Text;
               string yRange = YcoordinateAdrsBox.Text;
               string distanceRange = distAdrsBox.Text;
               
                if (okButtonActionEvent != null)
                    okButtonActionEvent.Invoke(xRange, yRange, distanceRange);
                this.Close();
            }
            else
            {
                MessageBox.Show("Mandatory field can't be left blank.");

            }
            
        }
        
        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            workSheet = Globals.ThisAddIn.Application.ActiveSheet;
            workSheet.SelectionChange += workSheet_SelectionChange;
            this.ActiveControl = XcoordinateAdrsBox;

        }

        //private TextBox lastFocused;
        
        void workSheet_SelectionChange(Microsoft.Office.Interop.Excel.Range Target)
        {
            
            if(textBoxSelector==1)
            {
                if(Target.Columns.Count==1)
                {
                    this.XcoordinateAdrsBox.Text = Target.Address;
                }
                else
                {

                }
            }
            else if(textBoxSelector==2)
            {
                if(Target.Columns.Count==1)
                {
                    this.YcoordinateAdrsBox.Text = Target.Address;
                }
                else
                {
                    //Target
                }
            }
            else if (textBoxSelector == 3)
            {
                if (Target.Columns.Count == 1)
                {
                    this.distAdrsBox.Text = Target.Address;
                }
                else
                {

                }
            }
            
        }

       
        private void XcoordinateAdrsBox_GotFocus(object sender, EventArgs e)
        {
            textBoxSelector = 1;
        }
        private void YcoordinateAdrsBox_GotFocus(object sender, EventArgs e)
        {
            textBoxSelector = 2;
        }
        private void distAdrsBox_GotFocus(object sender, EventArgs e)
        {
            textBoxSelector = 3; 
        }
        

        private void activWorkSheet_SelectionChange(Excel.Range Target)
        {
            throw new NotImplementedException();
        }

        protected override void OnClosed(EventArgs e)
        {
            base.OnClosed(e);
            workSheet.SelectionChange -= workSheet_SelectionChange;
        }
    }
}