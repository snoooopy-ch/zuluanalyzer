using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ZuluAnalyzer
{
    public partial class Setting : Form
    {
        public Setting()
        {
            InitializeComponent();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void Setting_Load(object sender, EventArgs e)
        {
            InitInterface();
        }

        private void InitInterface()
        {
            chkUseProxy.Checked = CGlobalVar.g_bUseProxy;
            txtProxyIP.Text = CGlobalVar.g_strProxyIP;
            txtProxyPort.Text = CGlobalVar.g_nProxyPort.ToString();
            txtProxyID.Text = CGlobalVar.g_strProxyID;
            txtProxyPassword.Text = CGlobalVar.g_strProxyPass;
            EnableProxyControls(CGlobalVar.g_bUseProxy);
        }

        private void EnableProxyControls(bool bEnable)
        {
            txtProxyIP.Enabled = bEnable;
            txtProxyPort.Enabled = bEnable;
            txtProxyID.Enabled = bEnable;
            txtProxyPassword.Enabled = bEnable;
        }


        private void btnOK_Click(object sender, EventArgs e)
        {
            CGlobalVar.g_bUseProxy = chkUseProxy.Checked;
            CGlobalVar.g_strProxyIP = txtProxyIP.Text;
            int.TryParse(txtProxyPort.Text, out CGlobalVar.g_nProxyPort);
            CGlobalVar.g_strProxyID = txtProxyID.Text;
            CGlobalVar.g_strProxyPass = txtProxyPassword.Text;

            CGlobalVar.WriteConfig();
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void chkUseProxy_CheckedChanged(object sender, EventArgs e)
        {
            EnableProxyControls(chkUseProxy.Checked);
        }

        private void btnProxyTest_Click(object sender, EventArgs e)
        {
            HttpCommon http_request = new HttpCommon();

            http_request.setURL("https://search.naver.com/search.naver?sm=top_hty&fbm=1&ie=utf8&query=%EB%82%B4+ip");
            http_request.setSendMode(HTTP_SEND_MODE.HTTP_GET);

            if (chkUseProxy.Checked)
            {
                string strProxyIP = txtProxyIP.Text;
                int nProxyPort = 0; int.TryParse(txtProxyPort.Text, out nProxyPort);
                string strProxyID = txtProxyID.Text;
                string strProxyPass = txtProxyPassword.Text;
                http_request.setProxy(strProxyIP, nProxyPort, strProxyID, strProxyPass);
            }

            if (!http_request.sendRequest(false, ""))
            {
                MessageBox.Show(Constant.PROXY_TEST_FAILED);
                return;
            }

            string response = http_request.getResponseString();
            int nStartPos = response.IndexOf("ip_chk_box");
            if (nStartPos < 0)
            {
                MessageBox.Show(Constant.PROXY_TEST_FAILED);
                return;
            }
            nStartPos = response.IndexOf("<em>", nStartPos);
            nStartPos += 4;
            int nEndPos = response.IndexOf("</em>", nStartPos);
            string strIp = response.Substring(nStartPos, nEndPos - nStartPos);

            MessageBox.Show("貴方のIPは「" + strIp + "」です。");
        }
    }
}
