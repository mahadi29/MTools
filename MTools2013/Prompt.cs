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
    
    public partial class Prompt: Form
    {
        public delegate void okButtonAction(string m1,string m2,string m3,string m4,string m5,string m6,string m7,string m8);
        public event okButtonAction okButtonActionEvent;
        
        //TextBox textBox;
        Excel.Worksheet workSheet;
        Excel.Worksheet activWorkSheet = Globals.ThisAddIn.GetActiveWorksheet();
        int textBoxSelector;

        public Prompt()
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
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.distLabel = new System.Windows.Forms.Label();
            this.rlLabel = new System.Windows.Forms.Label();
            this.xLabel = new System.Windows.Forms.Label();
            this.yLabel = new System.Windows.Forms.Label();
            this.riverNameTxtBox = new System.Windows.Forms.TextBox();
            this.topoIdTxtBox = new System.Windows.Forms.TextBox();
            this.SecIdAdrsBox = new System.Windows.Forms.TextBox();
            this.distAdrsBox = new System.Windows.Forms.TextBox();
            this.chainageAdrsBox = new System.Windows.Forms.TextBox();
            this.rlAdrsBox = new System.Windows.Forms.TextBox();
            this.XcoordinateAdrsBox = new System.Windows.Forms.TextBox();
            this.YcoordinateAdrsBox = new System.Windows.Forms.TextBox();
            this.okButton = new System.Windows.Forms.Button();
            this.cancelButton = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.closeButton = new System.Windows.Forms.Button();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(63, 64);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(66, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "River Name:";
            // 
            // distLabel
            // 
            this.distLabel.AutoSize = true;
            this.distLabel.Location = new System.Drawing.Point(237, 128);
            this.distLabel.Name = "distLabel";
            this.distLabel.Size = new System.Drawing.Size(52, 13);
            this.distLabel.TabIndex = 0;
            this.distLabel.Text = "Distance ";
            // 
            // rlLabel
            // 
            this.rlLabel.AutoSize = true;
            this.rlLabel.Location = new System.Drawing.Point(237, 153);
            this.rlLabel.Name = "rlLabel";
            this.rlLabel.Size = new System.Drawing.Size(80, 13);
            this.rlLabel.TabIndex = 0;
            this.rlLabel.Text = "Reduced Level";
            // 
            // xLabel
            // 
            this.xLabel.AutoSize = true;
            this.xLabel.Location = new System.Drawing.Point(153, 195);
            this.xLabel.Name = "xLabel";
            this.xLabel.Size = new System.Drawing.Size(14, 13);
            this.xLabel.TabIndex = 0;
            this.xLabel.Text = "X:";
            // 
            // distAdrsBox
            // 
            this.distAdrsBox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.distAdrsBox.Location = new System.Drawing.Point(318, 127);
            this.distAdrsBox.Name = "distAdrsBox";
            this.distAdrsBox.Margin = new Padding(5,3,0,0);
            this.distAdrsBox.AutoSize = false;
            this.distAdrsBox.Size = new System.Drawing.Size(68, 17);
            this.distAdrsBox.TabIndex = 1;
            this.distAdrsBox.GotFocus += new System.EventHandler(this.distAdrsBox_GotFocus);
            // 
            // rlAdrsBox
            // 
            this.rlAdrsBox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.rlAdrsBox.Location = new System.Drawing.Point(318, 152);
            this.rlAdrsBox.Name = "rlAdrsBox";
            this.rlAdrsBox.AutoSize = false;
            this.rlAdrsBox.Size = new System.Drawing.Size(68, 17);
            this.rlAdrsBox.TabIndex = 1;
            this.rlAdrsBox.GotFocus += new System.EventHandler(this.rlAdrsBox_GotFocus);
            // 
            // yLabel
            // 
            this.yLabel.AutoSize = true;
            this.yLabel.Location = new System.Drawing.Point(257, 195);
            this.yLabel.Name = "yLabel";
            this.yLabel.Size = new System.Drawing.Size(14, 13);
            this.yLabel.TabIndex = 0;
            this.yLabel.Text = "Y:";
            
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(240, 64);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(49, 13);
            this.label2.TabIndex = 0;
            this.label2.Text = "Topo ID:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(63, 128);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(57, 13);
            this.label3.TabIndex = 0;
            this.label3.Text = "Section ID";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(63, 153);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(52, 13);
            this.label4.TabIndex = 0;
            this.label4.Text = "Chainage";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(87, 195);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(58, 13);
            this.label5.TabIndex = 0;
            this.label5.Text = "Coordinate";
            // 
            // riverNameTxtBox
            // 
            this.riverNameTxtBox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.riverNameTxtBox.Location = new System.Drawing.Point(66, 79);
            this.riverNameTxtBox.Margin = new System.Windows.Forms.Padding(3,3,3,3);
            this.riverNameTxtBox.Name = "riverNameTxtBox";
            this.riverNameTxtBox.AutoSize = false;
            this.riverNameTxtBox.Size = new System.Drawing.Size(145, 17);
            this.riverNameTxtBox.TabIndex = 1;
            this.riverNameTxtBox.GotFocus += new System.EventHandler(this.riverNameTxtBox_GotFocus);
            // 
            // topoIdTxtBox
            // 
            this.topoIdTxtBox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.topoIdTxtBox.Location = new System.Drawing.Point(243, 79);
            this.topoIdTxtBox.Name = "topoIdTxtBox";
            this.topoIdTxtBox.AutoSize = false;
            this.topoIdTxtBox.Size = new System.Drawing.Size(145, 17);
            this.topoIdTxtBox.TabIndex = 1;
            this.topoIdTxtBox.GotFocus += new System.EventHandler(this.topoIdTxtBox_GotFocus);
            // 
            // SecIdAdrsBox
            // 
            this.SecIdAdrsBox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.SecIdAdrsBox.Location = new System.Drawing.Point(141, 127);
            this.SecIdAdrsBox.Name = "SecIdAdrsBox";
            this.SecIdAdrsBox.AutoSize = false;
            this.SecIdAdrsBox.Size = new System.Drawing.Size(70, 17);
            this.SecIdAdrsBox.TabIndex = 1;
            this.SecIdAdrsBox.GotFocus += new System.EventHandler(this.SecIdAdrsBox_GotFocus);
            // 
            // chainageAdrsBox
            // 
            this.chainageAdrsBox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.chainageAdrsBox.Location = new System.Drawing.Point(141, 152);
            this.chainageAdrsBox.Name = "chainageAdrsBox";
            this.chainageAdrsBox.AutoSize = false;
            this.chainageAdrsBox.Size = new System.Drawing.Size(70, 17);
            this.chainageAdrsBox.TabIndex = 1;
            this.chainageAdrsBox.GotFocus += new System.EventHandler(this.chainageAdrsBox_GotFocus);
            
            // 
            // XcoordinateAdrsBox
            // 
            this.XcoordinateAdrsBox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.XcoordinateAdrsBox.Location = new System.Drawing.Point(173, 194);
            this.XcoordinateAdrsBox.Name = "XcoordinateAdrsBox";
            this.XcoordinateAdrsBox.AutoSize = false;
            this.XcoordinateAdrsBox.Size = new System.Drawing.Size(70, 17);
            this.XcoordinateAdrsBox.TabIndex = 1;
            this.XcoordinateAdrsBox.GotFocus +=new System.EventHandler(this.XcoordinateAdrsBox_GotFocus);

            // 
            // YcoordinateAdrsBox
            // 
            this.YcoordinateAdrsBox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.YcoordinateAdrsBox.Location = new System.Drawing.Point(277, 194);
            this.YcoordinateAdrsBox.Name = "YcoordinateAdrsBox";
            this.YcoordinateAdrsBox.AutoSize = false;
            this.YcoordinateAdrsBox.Size = new System.Drawing.Size(70, 17);
            this.YcoordinateAdrsBox.TabIndex = 1;
            this.YcoordinateAdrsBox.GotFocus += new System.EventHandler(this.YcoordinateAdrsBox_GotFocus);

            // 
            // okButton
            // 
            this.okButton.BackColor = System.Drawing.Color.DeepSkyBlue;
            this.okButton.FlatAppearance.BorderSize = 0;
            this.okButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.okButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.okButton.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.okButton.Location = new System.Drawing.Point(80, 245);
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
            this.cancelButton.Location = new System.Drawing.Point(290, 245);
            this.cancelButton.Name = "cancelButton";
            this.cancelButton.Size = new System.Drawing.Size(80, 25);
            this.cancelButton.TabIndex = 2;
            this.cancelButton.Text = "Cancel";
            this.cancelButton.UseVisualStyleBackColor = false;
            this.cancelButton.Click += new System.EventHandler(this.cancelButton_Click);
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.SystemColors.ButtonFace; //.ControlDark; //.ButtonFace;//
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.riverNameTxtBox);
            this.panel1.Controls.Add(this.topoIdTxtBox);
            this.panel1.Controls.Add(this.SecIdAdrsBox);
            this.panel1.Controls.Add(this.distAdrsBox);
            this.panel1.Controls.Add(this.chainageAdrsBox);
            this.panel1.Controls.Add(this.rlAdrsBox);
            this.panel1.Controls.Add(this.XcoordinateAdrsBox);
            this.panel1.Controls.Add(this.YcoordinateAdrsBox);
            this.panel1.Controls.Add(this.okButton);
            this.panel1.Controls.Add(this.cancelButton);
            this.panel1.Controls.Add(this.closeButton);
            this.panel1.Controls.Add(this.label5);
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.xLabel);
            this.panel1.Controls.Add(this.yLabel);
            this.panel1.Controls.Add(this.rlLabel);
            this.panel1.Controls.Add(this.distLabel);
            this.panel1.Location = new System.Drawing.Point(1, 1);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(454, 299);
            this.panel1.TabIndex = 3;
            this.panel1.MouseDown += new System.Windows.Forms.MouseEventHandler(this.panel1_MouseDown);
            
            // 
            // closeButton
            // 
            this.closeButton.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(247)))), ((int)(((byte)(79)))), ((int)(((byte)(83)))));
            this.closeButton.FlatAppearance.BorderSize = 0;
            this.closeButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.closeButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.closeButton.Location = new System.Drawing.Point(423, -1);
            this.closeButton.Margin = new System.Windows.Forms.Padding(0);
            this.closeButton.Name = "closeButton";
            this.closeButton.Size = new System.Drawing.Size(30, 19);
            this.closeButton.TabIndex = 3;
            this.closeButton.Text = "X";
            this.closeButton.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            this.closeButton.UseVisualStyleBackColor = false;
            this.closeButton.Click += new System.EventHandler(this.closeButton_Click);
            // 
            // Prompt
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.ClientSize = new System.Drawing.Size(455, 300);
            this.Controls.Add(this.panel1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "Prompt";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Input Parameters";
            this.TopMost = true;
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);

        }
        
        private Label label1;
        private Label label2;
        private Label label3;
        private Label label4;
        private Label label5;
        private TextBox riverNameTxtBox;
        private TextBox topoIdTxtBox;
        private TextBox SecIdAdrsBox;
        private TextBox chainageAdrsBox;
        private TextBox XcoordinateAdrsBox;
        private TextBox YcoordinateAdrsBox;
        private Button okButton;
        private Button cancelButton;
        private Panel panel1;
        private Button closeButton;
        private TextBox rlAdrsBox;
        private TextBox distAdrsBox;
        private Label xLabel;
        private Label rlLabel;
        private Label distLabel;
        private Label yLabel;
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

            if(riverNameTxtBox.Text!="" & topoIdTxtBox.Text != "" & SecIdAdrsBox.Text!="" & distAdrsBox.Text != "" & chainageAdrsBox.Text != "" &rlAdrsBox.Text != "")
            {
               string riverName = riverNameTxtBox.Text;
               string topoID = topoIdTxtBox.Text;
               string secIdRange = SecIdAdrsBox.Text;
               string distanceRange = distAdrsBox.Text;
               string chainageRange = chainageAdrsBox.Text;
               string rlRange = rlAdrsBox.Text;
               string xRange = XcoordinateAdrsBox.Text;
               string yRange = YcoordinateAdrsBox.Text;

                if (okButtonActionEvent != null)
                    okButtonActionEvent.Invoke(riverName, topoID, secIdRange, distanceRange, chainageRange, rlRange, xRange, yRange);
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
            this.ActiveControl = riverNameTxtBox;

        }


        void workSheet_SelectionChange(Microsoft.Office.Interop.Excel.Range Target)
        {
            
            if(textBoxSelector==1)
            {
                if(Target.Columns.Count==1)
                {
                    this.SecIdAdrsBox.Text = Target.Address;
                }
                else
                {

                }
            }
            else if(textBoxSelector==2)
            {
                if(Target.Columns.Count==1)
                {
                    this.chainageAdrsBox.Text = Target.Address;
                }
                else
                {
                    //Target
                }
            }
            else if(textBoxSelector==3)
            {
                if (Target.Columns.Count == 1)
                {
                    this.XcoordinateAdrsBox.Text = Target.Address;
                }
                else
                {

                }
            }
            else if (textBoxSelector == 4)
            {
                if (Target.Columns.Count == 1)
                {
                    this.YcoordinateAdrsBox.Text = Target.Address;
                }
                else
                {

                }
            }
            else if (textBoxSelector == 5)
            {
                if (Target.Columns.Count == 1)
                {
                    this.distAdrsBox.Text = Target.Address;
                }
                else
                {

                }
            }
            else if (textBoxSelector == 6)
            {
                if (Target.Columns.Count == 1)
                {
                    this.rlAdrsBox.Text = Target.Address;
                }
                else
                {

                }
            }

        }

       
        private void SecIdAdrsBox_GotFocus(object sender, EventArgs e)
        {
            textBoxSelector = 1;
        }
        private void chainageAdrsBox_GotFocus(object sender, EventArgs e)
        {
            textBoxSelector = 2;
        }
        private void XcoordinateAdrsBox_GotFocus(object sender, EventArgs e)
        {
            textBoxSelector = 3;
        }
        private void YcoordinateAdrsBox_GotFocus(object sender, EventArgs e)
        {
            textBoxSelector = 4;
        }
        private void distAdrsBox_GotFocus(object sender, EventArgs e)
        {
            textBoxSelector = 5; 
        }
        private void rlAdrsBox_GotFocus(object sender, EventArgs e)
        {
            textBoxSelector = 6; 
        }

        private void riverNameTxtBox_GotFocus(object sender, EventArgs e)
        {
            textBoxSelector = 0;
        }
        private void topoIdTxtBox_GotFocus(object sender, EventArgs e)
        {
            textBoxSelector = 0;
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
