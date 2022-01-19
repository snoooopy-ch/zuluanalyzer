namespace ZuluAnalyzer
{
    partial class Setting
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Setting));
            this.tbcSetting = new System.Windows.Forms.TabControl();
            this.tpgBasic = new System.Windows.Forms.TabPage();
            this.groupBox7 = new System.Windows.Forms.GroupBox();
            this.chkUseProxy = new System.Windows.Forms.CheckBox();
            this.txtProxyPassword = new System.Windows.Forms.TextBox();
            this.label18 = new System.Windows.Forms.Label();
            this.txtProxyID = new System.Windows.Forms.TextBox();
            this.label21 = new System.Windows.Forms.Label();
            this.txtProxyPort = new System.Windows.Forms.TextBox();
            this.label19 = new System.Windows.Forms.Label();
            this.txtProxyIP = new System.Windows.Forms.TextBox();
            this.label20 = new System.Windows.Forms.Label();
            this.tpgPayment = new System.Windows.Forms.TabPage();
            this.btnOK = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnProxyTest = new System.Windows.Forms.Button();
            this.tbcSetting.SuspendLayout();
            this.tpgBasic.SuspendLayout();
            this.groupBox7.SuspendLayout();
            this.SuspendLayout();
            // 
            // tbcSetting
            // 
            this.tbcSetting.Controls.Add(this.tpgBasic);
            this.tbcSetting.Controls.Add(this.tpgPayment);
            this.tbcSetting.Location = new System.Drawing.Point(12, 12);
            this.tbcSetting.Name = "tbcSetting";
            this.tbcSetting.SelectedIndex = 0;
            this.tbcSetting.Size = new System.Drawing.Size(627, 362);
            this.tbcSetting.TabIndex = 0;
            // 
            // tpgBasic
            // 
            this.tpgBasic.BackColor = System.Drawing.Color.White;
            this.tpgBasic.Controls.Add(this.groupBox7);
            this.tpgBasic.ForeColor = System.Drawing.Color.White;
            this.tpgBasic.Location = new System.Drawing.Point(4, 22);
            this.tpgBasic.Name = "tpgBasic";
            this.tpgBasic.Padding = new System.Windows.Forms.Padding(3);
            this.tpgBasic.Size = new System.Drawing.Size(619, 336);
            this.tpgBasic.TabIndex = 0;
            this.tpgBasic.Text = "基本設定";
            // 
            // groupBox7
            // 
            this.groupBox7.Controls.Add(this.chkUseProxy);
            this.groupBox7.Controls.Add(this.txtProxyPassword);
            this.groupBox7.Controls.Add(this.label18);
            this.groupBox7.Controls.Add(this.txtProxyID);
            this.groupBox7.Controls.Add(this.label21);
            this.groupBox7.Controls.Add(this.txtProxyPort);
            this.groupBox7.Controls.Add(this.label19);
            this.groupBox7.Controls.Add(this.txtProxyIP);
            this.groupBox7.Controls.Add(this.label20);
            this.groupBox7.ForeColor = System.Drawing.Color.Black;
            this.groupBox7.Location = new System.Drawing.Point(13, 8);
            this.groupBox7.Name = "groupBox7";
            this.groupBox7.Size = new System.Drawing.Size(316, 163);
            this.groupBox7.TabIndex = 1;
            this.groupBox7.TabStop = false;
            this.groupBox7.Text = " プロキシ設定 ";
            // 
            // chkUseProxy
            // 
            this.chkUseProxy.AutoSize = true;
            this.chkUseProxy.Location = new System.Drawing.Point(22, 24);
            this.chkUseProxy.Name = "chkUseProxy";
            this.chkUseProxy.Size = new System.Drawing.Size(122, 17);
            this.chkUseProxy.TabIndex = 0;
            this.chkUseProxy.Text = "プロキシを使います。";
            this.chkUseProxy.UseVisualStyleBackColor = true;
            this.chkUseProxy.CheckedChanged += new System.EventHandler(this.chkUseProxy_CheckedChanged);
            // 
            // txtProxyPassword
            // 
            this.txtProxyPassword.Location = new System.Drawing.Point(105, 126);
            this.txtProxyPassword.Name = "txtProxyPassword";
            this.txtProxyPassword.PasswordChar = '*';
            this.txtProxyPassword.Size = new System.Drawing.Size(191, 20);
            this.txtProxyPassword.TabIndex = 8;
            // 
            // label18
            // 
            this.label18.AutoSize = true;
            this.label18.Location = new System.Drawing.Point(26, 129);
            this.label18.Name = "label18";
            this.label18.Size = new System.Drawing.Size(60, 13);
            this.label18.TabIndex = 7;
            this.label18.Text = "パスワード :";
            // 
            // txtProxyID
            // 
            this.txtProxyID.Location = new System.Drawing.Point(105, 100);
            this.txtProxyID.Name = "txtProxyID";
            this.txtProxyID.Size = new System.Drawing.Size(191, 20);
            this.txtProxyID.TabIndex = 6;
            // 
            // label21
            // 
            this.label21.AutoSize = true;
            this.label21.Location = new System.Drawing.Point(41, 103);
            this.label21.Name = "label21";
            this.label21.Size = new System.Drawing.Size(24, 13);
            this.label21.TabIndex = 5;
            this.label21.Text = "ID :";
            // 
            // txtProxyPort
            // 
            this.txtProxyPort.Location = new System.Drawing.Point(105, 73);
            this.txtProxyPort.Name = "txtProxyPort";
            this.txtProxyPort.Size = new System.Drawing.Size(191, 20);
            this.txtProxyPort.TabIndex = 4;
            // 
            // label19
            // 
            this.label19.AutoSize = true;
            this.label19.Location = new System.Drawing.Point(19, 76);
            this.label19.Name = "label19";
            this.label19.Size = new System.Drawing.Size(79, 13);
            this.label19.TabIndex = 3;
            this.label19.Text = "プロキシポート :";
            // 
            // txtProxyIP
            // 
            this.txtProxyIP.Location = new System.Drawing.Point(105, 47);
            this.txtProxyIP.Name = "txtProxyIP";
            this.txtProxyIP.Size = new System.Drawing.Size(191, 20);
            this.txtProxyIP.TabIndex = 2;
            // 
            // label20
            // 
            this.label20.AutoSize = true;
            this.label20.Location = new System.Drawing.Point(26, 50);
            this.label20.Name = "label20";
            this.label20.Size = new System.Drawing.Size(61, 13);
            this.label20.TabIndex = 1;
            this.label20.Text = "プロキシIP :";
            // 
            // tpgPayment
            // 
            this.tpgPayment.BackColor = System.Drawing.Color.White;
            this.tpgPayment.Location = new System.Drawing.Point(4, 22);
            this.tpgPayment.Name = "tpgPayment";
            this.tpgPayment.Size = new System.Drawing.Size(619, 438);
            this.tpgPayment.TabIndex = 1;
            this.tpgPayment.Text = "お支払方法";
            // 
            // btnOK
            // 
            this.btnOK.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(192)))));
            this.btnOK.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnOK.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
            this.btnOK.Location = new System.Drawing.Point(263, 380);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(111, 41);
            this.btnOK.TabIndex = 1;
            this.btnOK.Text = "設定する";
            this.btnOK.UseVisualStyleBackColor = false;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(192)))));
            this.btnCancel.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnCancel.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
            this.btnCancel.Location = new System.Drawing.Point(497, 380);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(111, 41);
            this.btnCancel.TabIndex = 2;
            this.btnCancel.Text = "キャンセル";
            this.btnCancel.UseVisualStyleBackColor = false;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnProxyTest
            // 
            this.btnProxyTest.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(192)))));
            this.btnProxyTest.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnProxyTest.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
            this.btnProxyTest.Location = new System.Drawing.Point(380, 380);
            this.btnProxyTest.Name = "btnProxyTest";
            this.btnProxyTest.Size = new System.Drawing.Size(111, 41);
            this.btnProxyTest.TabIndex = 3;
            this.btnProxyTest.Text = "プロキシテスト";
            this.btnProxyTest.UseVisualStyleBackColor = false;
            this.btnProxyTest.Click += new System.EventHandler(this.btnProxyTest_Click);
            // 
            // Setting
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(227)))), ((int)(((byte)(255)))), ((int)(((byte)(227)))));
            this.ClientSize = new System.Drawing.Size(651, 431);
            this.Controls.Add(this.btnProxyTest);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.tbcSetting);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "Setting";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "設定";
            this.Load += new System.EventHandler(this.Setting_Load);
            this.tbcSetting.ResumeLayout(false);
            this.tpgBasic.ResumeLayout(false);
            this.groupBox7.ResumeLayout(false);
            this.groupBox7.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl tbcSetting;
        private System.Windows.Forms.TabPage tpgBasic;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.GroupBox groupBox7;
        private System.Windows.Forms.TextBox txtProxyPassword;
        private System.Windows.Forms.Label label18;
        private System.Windows.Forms.TextBox txtProxyID;
        private System.Windows.Forms.Label label21;
        private System.Windows.Forms.TextBox txtProxyPort;
        private System.Windows.Forms.Label label19;
        private System.Windows.Forms.TextBox txtProxyIP;
        private System.Windows.Forms.Label label20;
        private System.Windows.Forms.CheckBox chkUseProxy;
        private System.Windows.Forms.Button btnProxyTest;
        private System.Windows.Forms.TabPage tpgPayment;
    }
}