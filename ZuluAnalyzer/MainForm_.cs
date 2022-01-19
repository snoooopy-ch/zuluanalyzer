using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json;
using System.Net.Http;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Diagnostics;
using System.Security.Cryptography;


namespace ZuluAnalyzer
{
    public delegate void callbackFormReadHistoryFile(bool result);
    public partial class MainForm : Form
    {
        private enum Result
        {
            FAILURE = 0,
            RETRY = 1,
            SUCCESS = 2,
        }

        private enum LogType
        {
            Download = 0,
            Combine = 1,
        }

        private class ComboItem
        {
            public int id { get; set; }
            public string name { get; set; }
        }

        /*
         * システムの組み合わせ状態を保管
         * spec:性能
         * strIndex:システムのIndex  
         */
        private class KumiawaseItem
        {
            public float spec { get; set; }
            public string strIndex { get; set; }
            
        }

        private HttpCommon http_request = new HttpCommon();

        delegate void AddTextLogCallBack(TextBox ctrl, string log);
        delegate void SetControlTextCallBack(Control ctrl, string text);
        delegate string GetControlTextCallBack(Control ctrl);

        int recentTradingPeriod = 0;
        int minimumTradingPeriod = 0;
        int minimunFollowingCount = 0;
        int downloadCount = 0;
        int currentDownload = 0;

        Object defaultArg = Type.Missing;
        ExchangeClass.ExchangeRate exchageRate = null;
        List<TradingSystem> systems = null;
        List<KumiawaseItem> linearList = null;
        List<KumiawaseItem> linearList2 = null;
        List<string> currentList = null;

        ListViewItem.ListViewSubItem SelectedLSI;

        int[] c = { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
        int[] c2 = { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };

        private string[] map = new string[]
        {
            "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"
        };

        public string getExcelColumnLetter(int number)
        {
            return map[number];
        }

        public MainForm()
        {
            InitializeComponent();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            CGlobalVar.ReadConfig();

            InitInterface();
            //DoLogin(Constant.FIRST_HTTP);
        }

        private void InitInterface()
        {
            txtRecentTradingPeriod.Text = CGlobalVar.g_nRecentTradingPeriod.ToString();
            txtMinimumTradingPeriod.Text = CGlobalVar.g_nMinimumTradingPeriod.ToString();
            txtMinimumFollowingCount.Text = CGlobalVar.g_nMinimumFollowingCount.ToString();
            txtDownloadCount.Text = CGlobalVar.g_nDownloadCount.ToString();

            txtLoginID.Text = CGlobalVar.g_strLoginID;
            txtPassword.Text = CGlobalVar.g_strPassword;

            chkUseProxy.Checked = CGlobalVar.g_bUseProxy;

            List<ComboItem> pp = new List<ComboItem>();
            pp.Add(new ComboItem() { id = -1, name = "ー選択ー" });
            pp.Add(new ComboItem() { id = 3, name = "3" });
            pp.Add(new ComboItem() { id = 5, name = "5" });
            pp.Add(new ComboItem() { id = 7, name = "7" });
            pp.Add(new ComboItem() { id = 9, name = "9" });
            pp.Add(new ComboItem() { id = 10, name = "10" });
            pp.Add(new ComboItem() { id = 13, name = "13" });
            pp.Add(new ComboItem() { id = 15, name = "15" });
            pp.Add(new ComboItem() { id = 17, name = "17" });
            pp.Add(new ComboItem() { id = 20, name = "20" });

            comboBox1.DisplayMember = "name";
            comboBox1.ValueMember = "id";
            comboBox1.DataSource = pp;
            comboBox1.SelectedIndex = 0;
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            CGlobalVar.WriteConfig();
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void txtRecentTradingPeriod_TextChanged(object sender, EventArgs e)
        {
            int.TryParse(txtRecentTradingPeriod.Text, out CGlobalVar.g_nRecentTradingPeriod);
        }

        private void txtMinimumTradingPeriod_TextChanged(object sender, EventArgs e)
        {
            int.TryParse(txtMinimumTradingPeriod.Text, out CGlobalVar.g_nMinimumTradingPeriod);
        }

        private void txtMinimumFollowingCount_TextChanged(object sender, EventArgs e)
        {
            int.TryParse(txtMinimumFollowingCount.Text, out CGlobalVar.g_nMinimumFollowingCount);
        }

        private void txtDownloadCount_TextChanged(object sender, EventArgs e)
        {
            int.TryParse(txtDownloadCount.Text, out CGlobalVar.g_nDownloadCount);
        }

        private void SetControlText(Control ctrl, string text)
        {
            if (ctrl == null) return;
            if (ctrl.InvokeRequired)
            {
                SetControlTextCallBack d = new SetControlTextCallBack(SetControlText);
                this.Invoke(d, new object[] { ctrl, text });
            }
            else
            {
                ctrl.Text = text;
            }
        }

        private string GetControlText(Control ctrl)
        {
            if (ctrl == null) return "";
            if (ctrl.InvokeRequired)
            {
                GetControlTextCallBack d = new GetControlTextCallBack(GetControlText);
                return (string)this.Invoke(d, new object[] { ctrl });
            }
            else
            {
                return ctrl.Text;
            }
        }

        private void btnClearLogs_Click(object sender, EventArgs e)
        {
            SetControlText(txtDLLogs, "");
        }

        private void btnSetting_Click(object sender, EventArgs e)
        {
            Setting setting = new Setting();
            if (setting.ShowDialog() == DialogResult.OK)
            {
                this.chkUseProxy.Checked = CGlobalVar.g_bUseProxy;
            }
        }

        private void txtLoginID_TextChanged(object sender, EventArgs e)
        {
            CGlobalVar.g_strLoginID = txtLoginID.Text;
        }

        private void txtPassword_TextChanged(object sender, EventArgs e)
        {
            CGlobalVar.g_strPassword = txtPassword.Text;
        }

        private void chkUseProxy_CheckStateChanged(object sender, EventArgs e)
        {
            CGlobalVar.g_bUseProxy = chkUseProxy.Checked;
        }

        private void AddTextLog(TextBox ctrl, string log)
        {
            string strFull = DateTime.Now.ToString("[yyyy-MM-dd HH:mm:ss:fff] ");

            if (ctrl.InvokeRequired)
            {
                AddTextLogCallBack d = new AddTextLogCallBack(AddTextLog);
                this.Invoke(d, new object[] { ctrl, log });
            }
            else
            {
                ctrl.Text += strFull + log;
                ctrl.SelectionStart = ctrl.Text.Length;
                ctrl.ScrollToCaret();
            }
        }

        private void AddLog(string log, LogType type = LogType.Download)
        {
            string strDate = DateTime.Now.ToString("yyyyMMdd");
            string strFull = DateTime.Now.ToString("[yyyy-MM-dd HH:mm:ss:fff] ");

            if (type == LogType.Download)
                AddTextLog(txtDLLogs, log + "\r\n");

            var string_buffer = new StringBuilder();
            string_buffer.Append(strFull + log);
            try
            {
                using (StreamWriter sw = File.AppendText("Logs/" + strDate + ".log"))
                {
                    sw.WriteLine(string_buffer.ToString());
                    sw.Close();
                }
            }
            catch (Exception ex)
            {
                string strErrMsg = "WriteLog Error " + ex.Message + "\n";
                if (type == LogType.Download)
                    AddTextLog(txtDLLogs, log + "\r\n");
            }
        }

        private void btnStartDownload_Click(object sender, EventArgs e)
        {
            btnStartDownload.Enabled = false;
            clearHistoryFiles();

            Thread thread = new Thread(() => loginAndDownload(Constant.OPENED_HTTP));
            thread.Start();
        }

        private void clearHistoryFiles()
        {
            var t = Directory.GetFiles("Downloads\\");
            var files = Directory.GetFiles("Downloads\\").Where(s => s.StartsWith("Downloads\\zulu_"));
            foreach (var file in files)
            {
                File.Delete(file);
            }
        }

        /*
         * ログインとスクレイピングを実行
         */
        private Result loginAndDownload(int flag)
        {
            currentDownload = 0;
            if (!checkInputValidate())
                return Result.RETRY;

            //相場データの取得
            if (!downloadExchangeRate())
                return Result.FAILURE;

            int nStartPos = 0, nEndPos = 0;
            try
            {
                if (flag == 0)
                    AddLog(Constant.SITE_LOGIN_STARTED);

                // Login - Step 1
                // ログインのヘーダを設定する。
                string url = Constant.g_strSiteLoginURL1;
                http_request.setURL(url);
                http_request.setSendMode(HTTP_SEND_MODE.HTTP_GET);
                http_request.setUserAgent("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/78.0.3904.97 Safari/537.36");
                http_request.appendCustomHeader("Accept-Language: ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7");
                http_request.appendCustomHeader("Accept-Encoding: gzip, deflate, br");
                http_request.appendCustomHeader("Cache-Control: max-age=0");
                http_request.appendCustomHeader("Upgrade-Insecure-Requests:	1");
                http_request.appendCustomHeader("Sec-Fetch-Site: cross-site");
                http_request.appendCustomHeader("Sec-Fetch-Mode: navigate");
                http_request.setAccept("text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3");
                http_request.setAutoRedirect(true);
                
                if (!http_request.sendRequest(false))
                {
                    if (flag == 1)
                        return Result.RETRY;

                    AddLog(Constant.SITE_LOGIN_FAILED);
                    return Result.RETRY;
                }

                AddLog(Constant.SITE_LOGIN_REQUEST_SUCCESS1);

                string response = http_request.getResponseString();
                if (response == "")
                {
                    AddLog(Constant.SITE_LOGIN_FAILED);
                    return Result.RETRY;
                }

                // Get form data
                nStartPos = response.IndexOf("id=\"__EVENTTARGET\" value=\"", 0);
                if (nStartPos < 0)
                {
                    AddLog(Constant.SITE_LOGIN_FAILED);
                    return Result.RETRY;
                }
                nStartPos += 26;
                nEndPos = response.IndexOf("\"", nStartPos);
                string strEventTarget = response.Substring(nStartPos, nEndPos - nStartPos);

                nStartPos = response.IndexOf("id=\"__EVENTARGUMENT\" value=\"", nEndPos);
                nStartPos += 28;
                nEndPos = response.IndexOf("\"", nStartPos);
                string strEventArgument = response.Substring(nStartPos, nEndPos - nStartPos);

                nStartPos = response.IndexOf("id=\"__CVIEWSTATE\" value=\"", nEndPos);
                nStartPos += 25;
                nEndPos = response.IndexOf("\"", nStartPos);
                string strCViewState = response.Substring(nStartPos, nEndPos - nStartPos);

                nStartPos = response.IndexOf("id=\"__VIEWSTATE\" value=\"", nEndPos);
                nStartPos += 24;
                nEndPos = response.IndexOf("\"", nStartPos);
                string strViewState = response.Substring(nStartPos, nEndPos - nStartPos);

                nStartPos = response.IndexOf("__RequestVerificationToken\" type=\"hidden\" value=\"", nEndPos);
                nStartPos += 49;
                nEndPos = response.IndexOf("\"", nStartPos);
                string strVerifyToken = response.Substring(nStartPos, nEndPos - nStartPos);

                // Login - Step 2
                var parameters = new Dictionary<string, object>
                {
                    { "__EVENTTARGET", strEventTarget },
                    { "__EVENTARGUMENT", strEventArgument },
                    { "__CVIEWSTATE", strCViewState },
                    { "__VIEWSTATE", strViewState },
                    { "__RequestVerificationToken", strVerifyToken },
                    { "ctl00$main$tbUsername", CGlobalVar.g_strLoginID },
                    { "ctl00$main$tbPassword", CGlobalVar.g_strPassword },
                    { "ctl00$main$btnLogin", "Login" },
                    { "ctl00$main$hdRedirect", "" },
                };

                var param = CGlobalVar.EncodeURIComponent(parameters);

                url = Constant.g_strSiteLoginURL2;
                http_request.setURL(url);
                http_request.setSendMode(HTTP_SEND_MODE.HTTP_POST);

                if (!http_request.sendRequest(false, param))
                {
                    AddLog(Constant.SITE_LOGIN_FAILED);
                    return Result.RETRY;
                }

                AddLog(Constant.SITE_LOGIN_REQUEST_SUCCESS2);

                response = http_request.getResponseString();
                if (response == "" || (!response.Contains("資産") && !response.Contains("Equity")))
                {
                    AddLog(Constant.SITE_LOGIN_FAILED);
                    return Result.RETRY;
                }

                // Login - Step 3
                url = Constant.g_strSiteTradeURL;
                http_request.setURL(url);
                http_request.setSendMode(HTTP_SEND_MODE.HTTP_GET);
                http_request.setAutoRedirect(true);

                if (!http_request.sendRequest(false))
                {
                    AddLog(Constant.SITE_LOGIN_FAILED);
                    return Result.RETRY;
                }

                url = Constant.g_strSiteGetToken;
                http_request.setURL(url);
                http_request.setSendMode(HTTP_SEND_MODE.HTTP_POST);
                http_request.setAutoRedirect(true);
                if (!http_request.sendRequest(false))
                {
                    AddLog(Constant.SITE_LOGIN_FAILED);
                    return Result.RETRY;
                }

                response = http_request.getResponseString();
                var responseObj = JsonConvert.DeserializeObject<JWTResponse>(response);

                string strToken = responseObj.accesToken;
                if (strToken.Equals(""))
                {
                    AddLog(Constant.SITE_LOGIN_FAILED);
                    return Result.RETRY;
                }

                int page = 0;
                AddLog(Constant.SITE_TRADER_DOWNLOADING);
                TraderBaseClass.TraderResponse allTraders = new TraderBaseClass.TraderResponse();

                /*
                 * トレードの基礎資料を保管する。
                 */
                do
                {
                    url = Constant.g_strSiteTrader;
                    http_request.setURL(url);
                    http_request.setSendMode(HTTP_SEND_MODE.HTTP_POST);
                    http_request.setAutoRedirect(true);
                    http_request.appendCustomHeader("authorization: Bearer " + strToken);

                    /*
                     * トレードのリストリクエストのヘッダー
                     */
                    parameters = new Dictionary<string, object>
                    {
                        { "flavor", "global" },
                        { "minPips", 0.1 },
                        { "page", page },
                        { "size", 100 },
                        { "sortAsc", false },
                        { "sortBy", "liveFollowersProfit" },
                        { "timeFrame", 30 },
                    };

                    var serizedString = JsonConvert.SerializeObject(parameters);
                    if (!http_request.sendRequest(false, serizedString, "application/json"))
                    {
                        AddLog(Constant.SITE_LOGIN_FAILED);
                        return Result.RETRY;
                    }

                    response = http_request.getResponseString();

                    var traders = JsonConvert.DeserializeObject<TraderBaseClass.TraderResponse>(response);
                    if (traders.result.Count == 0)
                    {
                        break;
                    }

                    if (allTraders.result == null)
                        allTraders.result = traders.result;
                    else
                        allTraders.result.AddRange(traders.result);
                    page++;

                    Invoke((MethodInvoker)(() => textTrader.Text = allTraders.result.Count.ToString()));
                }
                while (true);

                if (allTraders.result.Count == 0)
                {
                    AddLog(Constant.SITE_FETCH_TRADER_FAILED);
                    return Result.RETRY;
                }

                AddLog(Constant.SITE_FETCH_TRADER_SUCCESS);

                int checkCount = 1;

                //別々トレードの取引履歴リクエスト
                foreach (var trader in allTraders.result)
                {

                    Invoke((MethodInvoker)(() => txtDLFinishedCount.Text = string.Format("{0}人中{1}人検査、成功：{2}", allTraders.result.Count, checkCount, currentDownload)));

                    if (checkCount == 390)
                    {
                        int j = 0;
                        j += 1;
                    }
                    checkCount++;

                    url = string.Format(Constant.g_strSiteTraderDetail, trader.trader.providerId);
                    http_request.setURL(url);
                    http_request.setSendMode(HTTP_SEND_MODE.HTTP_GET);
                    http_request.setAutoRedirect(true);

                    if (checkCount == 300)
                        Thread.Sleep(2000);
                    if (!http_request.sendRequest(false))
                    {
                        /*
                         * リクエストに失敗すればヘッダーを再度設定。
                         */
                        http_request.setUserAgent("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/78.0.3904.97 Safari/537.36");
                        http_request.appendCustomHeader("Accept-Language: ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7");
                        http_request.appendCustomHeader("Accept-Encoding: gzip, deflate, br");
                        http_request.appendCustomHeader("Cache-Control: max-age=0");
                        http_request.appendCustomHeader("Upgrade-Insecure-Requests:	1");
                        http_request.appendCustomHeader("Sec-Fetch-Site: cross-site");
                        http_request.appendCustomHeader("Sec-Fetch-Mode: navigate");
                        http_request.appendCustomHeader("authorization: Bearer " + strToken);
                        http_request.setAccept("text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3");
                        AddLog(string.Format(Constant.SITE_FETCH_TRADEHISTORY_FAILED, trader.trader.profile.name, trader.trader.providerId));
                        continue;
                    }

                    response = http_request.getResponseString();
                    var history = JsonConvert.DeserializeObject<TradeHistoryClass.TraderContent>(response);
                    long firstDateOpen = history.content[0].dateOpen;

                    if (firstDateOpen == 0)
                        continue;

                    trader.trader.overallStats.firstOpenTradeDate = firstDateOpen;
                    
                    //トレードがダウンロードの条件に合うかを検査。
                    if (!checkTraderWithCondition(trader))
                    {
                        continue;
                    }

                    url = string.Format(Constant.g_strSiteExportxls, trader.trader.providerId);
                    http_request.setURL(url);
                    http_request.setSendMode(HTTP_SEND_MODE.HTTP_GET);
                    http_request.setAutoRedirect(true);

                    string file = "zulu_" + cleanTraderName(trader.trader.profile.name) + "_" + trader.trader.providerId + 
                        "_" + trader.trader.profile.baseCurrencyName + "(" +trader.trader.profile.baseCurrencySymbol + ").xlsx";
                    FileStream fs = File.OpenWrite("Downloads/" + file);
                    if (!http_request.sendRequest(false, "", "image/jpg", fs))
                    {
                        AddLog(string.Format(Constant.SITE_EXPORT_XLS_FAILED, trader.trader.profile.name, trader.trader.providerId));
                        continue;
                    }
                    fs.Close();

                    addListView(trader, file);
                    currentDownload++;

                    Invoke((MethodInvoker)(() => txtDLFinishedCount.Text = string.Format("{0}人中{1}人検査、成功：{2}", allTraders.result.Count, checkCount, currentDownload)));
                    if (downloadCount <= currentDownload)
                        break;
                }

                AddLog(string.Format(Constant.SITE_FATCH_HISTORY_COUNT, currentDownload));
                
                return Result.SUCCESS;
            }
            catch (Exception ex)
            {
                string strErrMsg = Constant.SITE_LOGIN_FAILED + " " + Constant.ERROR_MESSAGE + ex.ToString();
                AddLog(strErrMsg);

                return Result.FAILURE;
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();

            Invoke((MethodInvoker)(() => btnStartDownload.Enabled = true));

        }

        private void addListView(TraderBaseClass.Trader trader, string file)
        {
            ListViewItem lvi = new ListViewItem(file);
            lvi.SubItems.Add(trader.trader.profile.name);
            lvi.SubItems.Add(trader.trader.timeframeStats.i30.trades.ToString());
            lvi.SubItems.Add(trader.trader.timeframeStats.i30.overallDrawDown.ToString());
            lvi.SubItems.Add(trader.trader.timeframeStats.i30.avgPipsPerTrade.ToString());
            lvi.SubItems.Add(trader.trader.overallStats.followers.ToString());
            lvi.SubItems.Add(trader.trader.profile.baseCurrencyName);
            lvi.Checked = true;
            lvi.UseItemStyleForSubItems = false;
            if (currentList.Contains(trader.trader.profile.baseCurrencyName + "/JPY"))
            {
                Color color = GetColor(trader.trader.profile.baseCurrencyName);
                lvi.SubItems[6].BackColor = color;
            }

            Invoke((MethodInvoker)(() => fileListView.Items.Add(lvi)));
        }

        private bool checkInputValidate()
        {
            if (!int.TryParse(txtRecentTradingPeriod.Text, out recentTradingPeriod))
            {
                MessageBox.Show(Constant.ERROR_MESSAGE_RECENT_PERIOD);
                return false;
            }

            if (!int.TryParse(txtMinimumTradingPeriod.Text, out minimumTradingPeriod))
            {
                MessageBox.Show(Constant.ERROR_MESSAGE_MIN_FOLLOWER);
                return false;
            }

            if (!int.TryParse(txtMinimumFollowingCount.Text, out minimunFollowingCount))
            {
                MessageBox.Show(Constant.ERROR_MESSAGE_MIN_FOLLOWER);
                return false;
            }

            if (!int.TryParse(txtDownloadCount.Text, out downloadCount))
            {
                MessageBox.Show(Constant.ERROR_MESSAGE_DOWNLOAD);
                return false;
            }
            return true;
        }

        private bool checkTraderWithCondition(TraderBaseClass.Trader trader)
        {
            string blackerText = textBoxBlacker.Text;
            string[] blackers = blackerText.Split(';');

            //ブラックリストの検査
            if (!(blackers.Count() == 1 && blackers[0].Equals("")) && blackers.Any(trader.trader.profile.name.Contains))
            {
                AddLog(string.Format(Constant.ERROR_MESSAGE_BLOCKED, trader.trader.profile.name));
                return false;
            }

            //最低取引期間の検査
            double d = double.Parse(trader.trader.overallStats.lastOpenTradeDate.ToString());
            TimeSpan ts = TimeSpan.FromMilliseconds(d);
            DateTime lastOp = new DateTime(1970, 1, 1) + ts;

            d = double.Parse(trader.trader.overallStats.firstOpenTradeDate.ToString());
            ts = TimeSpan.FromMilliseconds(d);
            DateTime firstOp = new DateTime(1970, 1, 1) + ts;

            int follower = trader.trader.overallStats.followers;
            if (minimunFollowingCount != 0)
            {
                if (minimunFollowingCount > follower)
                    return false;
            }
            
            double t;
            //最近取引期間の検査
            if (minimumTradingPeriod != 0)
            {
                long tradePeriod = trader.trader.overallStats.lastOpenTradeDate - trader.trader.overallStats.firstOpenTradeDate;
                t = (double)tradePeriod / 1000.0 / 60.0 / 60.0 / 24.0 / 30.0;
                if (minimumTradingPeriod > t)
                    return false;
            }
            
            if (recentTradingPeriod != 0)
            {
                TimeSpan span = DateTime.Now - (new DateTime(1970, 1, 1));
                double now = span.TotalMilliseconds;
                now = now - trader.trader.overallStats.lastOpenTradeDate;
                t = (double)now / 1000.0 / 60.0 / 60.0 / 24.0 / 30.0;
                if (recentTradingPeriod < t)
                    return false;
            }
            return true;
        }

        private float getRate(string currentName)
        {
            string pair = currentName + "/JPY";
            foreach(ExchangeClass.ExchangeTable table in exchageRate.forexList)
            {
                if (table.ticker.Equals(pair))
                {
                    return table.open;
                }
            }

            return 1;
        }

        private bool contentShiftRightExcel()
        {
            foreach (ListViewItem lvi in fileListView.Items)
            {
                try
                {
                    if (!lvi.Checked)
                        continue;

                    string fileName = "Downloads//" + lvi.Text;
                    string validFileName = Path.GetFileNameWithoutExtension(fileName);
                    string[] segs = validFileName.Split('_');
                    string traderName = "", traderId, baseCurrencyNameSymbol = "";

                    /*
                     * トレードの名前・ID・通貨
                     */
                    if (segs.Length >= 4)
                    {
                        int k;
                        for (k = 1; k < segs.Length - 3; k++)
                            traderName += (segs[k] + "_");
                        traderName += segs[k];
                        traderId = segs[k + 1];
                        baseCurrencyNameSymbol = segs[k + 2];
                    }

                    string[] arr = baseCurrencyNameSymbol.Split(new char[] { '(', ')' }, StringSplitOptions.None);
                    string baseCurrencySymbol = arr[1];
                    string baseCurrencyName = baseCurrencyNameSymbol.Substring(0, baseCurrencyNameSymbol.IndexOf('(')); ;
                    fileName = System.IO.Path.GetFullPath(fileName);

                    Excel.Application xlApp = new Excel.Application();
                    xlApp.Visible = false;
                    Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(fileName, 0, false, 5, "", "", false,
                        Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                    Excel.Worksheet worksheet = (Excel.Worksheet)xlWorkBook.ActiveSheet;

                    // If already shifted, continue;
                    Excel.Range oRng = worksheet.Range["A1"];
                    if (oRng.Value2 != "System ID")
                    {
                        oRng = worksheet.Range["A1"];
                        oRng.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight,
                            Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow);
                        oRng = worksheet.Range["A1"];
                        oRng.Value2 = "System ID";

                        // Find the last real row
                        int totalRows = worksheet.Cells.Find("*", System.Reflection.Missing.Value,
                                           System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                                           Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                                           false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

                        // Find the last real column
                        int totalColumns = worksheet.Cells.Find("*", System.Reflection.Missing.Value,
                                                       System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                                                       Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlPrevious,
                                                       false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Column;

                        float rate = getRate(baseCurrencyName);
                        
                        if (rate == 1)
                        {
                            AddLog(string.Format(Constant.ZULUANALYSE_EXCHAGE_FAILED, baseCurrencyName));
                            lvi.Checked = false;
                            continue;
                        }

                        oRng = worksheet.Range["A2", "A" + totalRows];
                        oRng.Value2 = traderName;
                        oRng.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightYellow);

                        string multiplyer = string.Format("=O2*{0}", rate);
                        oRng = worksheet.Range["P2", "P" + totalRows];
                        oRng.Formula = multiplyer;
                        oRng.NumberFormat = "¥#;[Red]-¥#";

                        oRng = worksheet.Range["P1"];
                        string profit = "Profit(¥)";
                        oRng.Value2 = profit;

                        oRng = worksheet.Range["O1"];
                        profit = "Profit(" + baseCurrencySymbol + ")";
                        oRng.Value2 = profit;
                        worksheet.Cells[1, 16].Font.FontStyle = "Bold";

                        oRng = worksheet.Range["M1"];
                        profit = "Interest(" + baseCurrencySymbol + ")";
                        oRng.Value2 = profit;

                        worksheet.Calculate();
                    }
                        
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRng);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                    xlWorkBook.Close(true);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkBook);
                    xlApp.UserControl = true;
                    xlApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject((object)xlApp);

                    Invoke((MethodInvoker)(() => progressBar1.Value = progressBar1.Value + 1));
                    Invoke((MethodInvoker)(() => AddLog(String.Format("{0}の履歴処理１段階の成功", lvi.Text))));
                }
                catch (Exception e)
                {
                    Invoke((MethodInvoker)(() => AddLog(String.Format("{0}の履歴を処理１段階の失敗", lvi.Text))));
                    continue;

                }
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();

            return true;
        }

        /*
         * トレードの履歴ファイルの結合
         */
        private bool mergeAllHistory(int start)
        {
            int current = 0;
            bool retval = true;
                        
            int currentRowDst = 0;
            int currentColDst = 0;
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkBookDst = null;
            Excel.Worksheet worksheetDst = null;
            Excel.Workbook xlWorkBookSrc = null;
            Excel.Worksheet worksheetSrc = null;
            xlApp.Visible = false;

            string dstFileName = "Downloads//All_履歴.xlsx";

            foreach (ListViewItem lvi in fileListView.Items)
            {
                current++;
                if (current < start)
                {
                    continue;
                }

                if (current == start + 50)
                {
                    break;
                }
                try
                {
                    string srcFileName = "Downloads//" + lvi.Text;
                    if (!lvi.Checked)
                        continue;

                    if (!File.Exists(dstFileName))
                    {
                        /*
                         * ファイルがない場合は最初のファイルをコピーします。
                         */
                        System.IO.File.Copy(srcFileName, dstFileName, true);

                        dstFileName = System.IO.Path.GetFullPath(dstFileName);
                        xlWorkBookDst = xlApp.Workbooks.Open(dstFileName, 0, false, 5, "", "", false,
                            Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                        worksheetDst = (Excel.Worksheet)xlWorkBookDst.ActiveSheet;

                        currentRowDst = worksheetDst.Cells.Find("*", System.Reflection.Missing.Value,
                                       System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                                       Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                                       false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

                        currentColDst = worksheetDst.Cells.Find("*", System.Reflection.Missing.Value,
                                       System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                                       Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                                       false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Column;
                    }
                    else
                    {
                        /*
                         * ファイルがある場合はファイルの内容だけコピーします。
                         */
                        srcFileName = System.IO.Path.GetFullPath(srcFileName);
                        dstFileName = System.IO.Path.GetFullPath(dstFileName);
                        xlWorkBookSrc = xlApp.Workbooks.Open(srcFileName, 0, false, 5, "", "", false,
                            Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                        xlWorkBookDst = xlApp.Workbooks.Open(dstFileName, 0, false, 5, "", "", false,
                            Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                        worksheetSrc = (Excel.Worksheet)xlWorkBookSrc.ActiveSheet;
                        worksheetDst = (Excel.Worksheet)xlWorkBookDst.ActiveSheet;

                        int limit = 1000000;
                        int totalRowsSrc = worksheetSrc.Cells.Find("*", System.Reflection.Missing.Value,
                                       System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                                       Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                                       false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

                        currentRowDst = worksheetDst.Cells.Find("*", System.Reflection.Missing.Value,
                                       System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                                       Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                                       false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;


                        if ((currentRowDst + totalRowsSrc) <= limit)
                        {
                            int sheetColCount = worksheetSrc.Cells.Find("*", System.Reflection.Missing.Value,
                                           System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                                           Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlPrevious,
                                           false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Column;

                            string str = getExcelColumnLetter(sheetColCount - 1);
                            Excel.Range range1 = worksheetSrc.get_Range("A2", str + totalRowsSrc.ToString());
                            range1.Copy(worksheetDst.get_Range(string.Format("A{0}", currentRowDst + 1), Missing.Value));

                            currentRowDst += (totalRowsSrc - 1);
                        }
                        
                        xlWorkBookDst.Save();

                        worksheetDst.Name = "All_履歴";

                        System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheetSrc);
                        xlWorkBookSrc.Close(true);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkBookSrc);

                        System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheetDst);
                        xlWorkBookDst.Close(true);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkBookDst);

                        Invoke((MethodInvoker)(() => AddLog(String.Format("{0}の履歴処理２段階の成功", lvi.Text))));
                    }
                }
                catch (Exception e)
                {
                    AddLog(e.Message);
                    Invoke((MethodInvoker)(() => AddLog(String.Format("{0}の履歴処理２段階の失敗", lvi.Text))));
                }
                finally
                {
                    
                }
            }
            xlApp.UserControl = true;
            xlApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject((object)xlApp);

            GC.Collect();
            GC.WaitForPendingFinalizers();

            return retval;
        }

        private bool deleteColumn()
        {
            bool retval = true;

            try
            {
                int currentRowDst = 0;
                int currentColDst = 0;
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkBookDst = null;
                Excel.Worksheet worksheetDst = null;
                xlApp.Visible = false;

                string dstFileName = "Downloads//All_履歴.xlsx";

                dstFileName = System.IO.Path.GetFullPath(dstFileName);
                xlWorkBookDst = xlApp.Workbooks.Open(dstFileName, 0, false, 5, "", "", false,
                    Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                worksheetDst = (Excel.Worksheet)xlWorkBookDst.ActiveSheet;

                currentRowDst = worksheetDst.Cells.Find("*", System.Reflection.Missing.Value,
                                System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                                Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                                false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

                currentColDst = worksheetDst.Cells.Find("*", System.Reflection.Missing.Value,
                                System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                                Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                                false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Column;

                
                int limit = 1000000;
                int sheetColCount = worksheetDst.Cells.Find("*", System.Reflection.Missing.Value,
                                System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                                Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlPrevious,
                                false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Column;

                string str = getExcelColumnLetter(sheetColCount - 1);
                Excel.Range range1 = worksheetDst.get_Range("P1", "P" + currentRowDst.ToString());
                range1.Copy();
                worksheetDst.get_Range("O1", "O" + currentRowDst.ToString()).PasteSpecial(Excel.XlPasteType.xlPasteValues, 
                    Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                worksheetDst.get_Range("O1", "O" + currentRowDst.ToString()).NumberFormat = "¥#;[Red]-¥#";
                range1.Clear();
                worksheetDst.Calculate();

                System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheetDst);
                xlWorkBookDst.Close(true);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkBookDst);
                xlApp.UserControl = true;
                xlApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject((object)xlApp);

                progressBar1.Value += 1;
            }
            catch (Exception e)
            {
                AddLog(e.Message);
                retval = false;
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            return retval;
        }

        private bool zuluAnalyzerStep1()
        {
            bool retval = true;
            string dstFileName = "Downloads//All_履歴.xlsx";

            try
            {
                int currentRowSrc = 0;
                int currentColSrc = 0;
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkBookDst = null;
                Excel.Worksheet worksheetDst = null;
                Excel.Worksheet worksheetSrc = null;
                xlApp.Visible = false;
                dstFileName = System.IO.Path.GetFullPath(dstFileName);

                xlWorkBookDst = xlApp.Workbooks.Open(dstFileName, 0, false, 5, "", "", false,
                    Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                
                worksheetSrc = (Excel.Worksheet)xlWorkBookDst.ActiveSheet;
                worksheetDst = (Excel.Worksheet)xlWorkBookDst.Worksheets.Add(defaultArg, worksheetSrc, defaultArg, defaultArg);
                worksheetDst.Name = "②計算資料";

                currentRowSrc = worksheetSrc.Cells.Find("*", System.Reflection.Missing.Value,
                               System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                               Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                               false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

                currentColSrc = worksheetSrc.Cells.Find("*", System.Reflection.Missing.Value,
                               System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                               Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                               false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Column;


                string str = getExcelColumnLetter(currentColSrc - 1);
                Excel.Range range = worksheetSrc.get_Range("A1", str + currentRowSrc.ToString());
                range.Copy(worksheetDst.get_Range("A1", Missing.Value));

                worksheetDst.Cells[1, 17].Value = "計算行①";
                Excel.Range r1 = worksheetDst.get_Range("Q2", "Q" + currentRowSrc.ToString());
                r1.Formula = "=IF(N2*F2=0,0,ABS((O2)/N2/F2))*0.01";

                worksheetDst.Cells[1, 18].value = "計算行②Highest Profit(¥)";
                Excel.Range r2 = worksheetDst.get_Range("R2", "R" + currentRowSrc.ToString());
                r2.Formula = "=IF(VALUE(TRIM(CLEAN(F2)))=0,0,K2*Q2/F2*0.01)";
                r2.Cells.NumberFormat = "#.00;[Red]-¥#.00";

                worksheetDst.Cells[1, 19].value = "計算行  ③Worst Drawdown(¥)";
                Excel.Range r3 = worksheetDst.get_Range("S2", "S" + currentRowSrc.ToString());
                r3.Formula = "=IF(VALUE(TRIM(CLEAN(F2)))=0,0,L2*Q2/F2*0.01)";
                r3.Cells.NumberFormat = "#.00;[Red]-¥#.00";
                
                Excel.Range r4 = worksheetDst.get_Range("G2", "G" + currentRowSrc.ToString());
                r4.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightYellow);
                
                Excel.Range r5 = worksheetDst.get_Range("H2", "H" + currentRowSrc.ToString());
                r5.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightYellow);
                
                worksheetDst.Cells[1, 11].value = "Highest Profit(¥)";
                worksheetDst.Cells[1, 12].value = "Worst Drawdown(¥)";
                worksheetDst.Calculate();
                xlWorkBookDst.Save();

                Excel.Range r6 = worksheetDst.get_Range("Y2", "Y" + currentRowSrc.ToString());
                r6.Formula = "=LEFT(G2, 10)";

                Excel.Range r7 = worksheetDst.get_Range("Z2", "Z" + currentRowSrc.ToString());
                r7.Formula = "=LEFT(H2, 10)";
                worksheetDst.Calculate();
                xlWorkBookDst.Save();

                xlApp.DisplayAlerts = false;
                Excel.Range rg = worksheetDst.get_Range("Y2", "Y" + currentRowSrc.ToString());
                rg.Copy();
                worksheetDst.get_Range("G2", "G" + currentRowSrc.ToString()).PasteSpecial(Excel.XlPasteType.xlPasteValues,
                    Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                rg = worksheetDst.get_Range("Z2", "Z" + currentRowSrc.ToString());
                rg.Copy();
                worksheetDst.get_Range("H2", "H" + currentRowSrc.ToString()).PasteSpecial(Excel.XlPasteType.xlPasteValues,
                    Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

                r6.Clear();
                r7.Clear();

                //Calculate the formula for whole worksheet
                System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheetSrc);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheetDst);
                xlWorkBookDst.Close(true);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkBookDst);
                xlApp.UserControl = true;
                xlApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject((object)xlApp);

                AddLog(Constant.ZULUANALYSE_STEP1_SUCCESS);
            }
            catch (Exception e)
            {
                AddLog(Constant.ZULUANALYSE_STEP1_FAILED);
                AddLog(e.Message);
                retval = false;
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            return retval;
        }

        private bool zuluAnalyzerStep2()
        {
            bool retval = true;
            string dstFileName = "Downloads//All_履歴.xlsx";

            try
            {
                int currentRowSrc = 0;
                int currentColSrc = 0;
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkBookDst = null;
                Excel.Worksheet worksheetDst = null;
                Excel.Worksheet worksheetSrc = null;
                xlApp.Visible = false;
                dstFileName = System.IO.Path.GetFullPath(dstFileName);

                xlWorkBookDst = xlApp.Workbooks.Open(dstFileName, 0, false, 5, "", "", false,
                    Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

                worksheetSrc = (Excel.Worksheet)xlWorkBookDst.Worksheets["②計算資料"];
                worksheetDst = (Excel.Worksheet)xlWorkBookDst.Worksheets.Add(defaultArg, worksheetSrc, defaultArg, defaultArg);
                worksheetDst.Name = "③計算結果代入+DIV 0!項目削除, ⑤ロット数計算結果代入";

                currentRowSrc = worksheetSrc.Cells.Find("*", System.Reflection.Missing.Value,
                               System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                               Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                               false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

                currentColSrc = worksheetSrc.Cells.Find("*", System.Reflection.Missing.Value,
                               System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                               Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                               false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Column;

                //データのコピー
                string str = getExcelColumnLetter(currentColSrc - 1);
                Excel.Range range = worksheetSrc.get_Range("A1", str + currentRowSrc.ToString());
                range.Copy(worksheetDst.get_Range("A1", Missing.Value));
                xlApp.DisplayAlerts = false;

                // EXLSの作成
                Excel.Range rg = worksheetDst.get_Range("R2", "R" + currentRowSrc.ToString());
                rg.Copy();
                Excel.Range rg1 = worksheetDst.get_Range("K2", "K" + currentRowSrc.ToString());
                rg1.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                rg1.EntireColumn.NumberFormat = "¥#";
                rg.Cells.Value = "1";

                // 19 --> 12
                rg = worksheetDst.get_Range("S2", "S" + currentRowSrc.ToString());
                rg.Copy();
                rg1 = worksheetDst.get_Range("L2", "L" + currentRowSrc.ToString());
                rg1.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                rg1.EntireColumn.NumberFormat = "¥#";
                rg.Cells.Clear();

                rg = worksheetDst.get_Range("F2", "F" + currentRowSrc.ToString());
                rg.Cells.Value = 0.01;

                Excel.Range r20 = worksheetDst.get_Range("T2", "T" + currentRowSrc.ToString());
                r20.Formula = "=LEFT(G2, 4) & \"年\"";
                Excel.Range r21 = worksheetDst.get_Range("U2", "U" + currentRowSrc.ToString());
                r21.Formula = "=MID(G2, 6, 2) & \"月\"";

                rg = worksheetDst.get_Range("T2", "T" + currentRowSrc.ToString());
                rg.Copy();
                rg1 = worksheetDst.get_Range("P2", "P" + currentRowSrc.ToString());
                rg1.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                rg.Cells.Clear();

                rg = worksheetDst.get_Range("U2", "U" + currentRowSrc.ToString());
                rg.Copy();
                rg1 = worksheetDst.get_Range("Q2", "Q" + currentRowSrc.ToString());
                rg1.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                rg.Cells.Clear();

                worksheetDst.Cells[1, 16].value = "Year";
                worksheetDst.Cells[1, 17].value = "Month";
                worksheetDst.Cells[1, 18].value = "回数";
                worksheetDst.Cells[1, 16].Font.FontStyle = "Bold";
                worksheetDst.Cells[1, 17].Font.FontStyle = "Bold";
                worksheetDst.Cells[1, 18].Font.FontStyle = "Bold";
                worksheetDst.Cells[1, 19].value = "";

                Excel.Range r1 = worksheetDst.get_Range("K2", "K" + currentRowSrc.ToString());
                r1.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightYellow);

                Excel.Range r2 = worksheetDst.get_Range("L2", "L" + currentRowSrc.ToString());
                r2.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightYellow);
                r2.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);

                r1 = worksheetDst.get_Range("P2", "P" + currentRowSrc.ToString());
                r1.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightYellow);

                r1 = worksheetDst.get_Range("Q2", "Q" + currentRowSrc.ToString());
                r1.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightYellow);

                //Calculate the formula for whole worksheet
                worksheetDst.Calculate();
                worksheetDst.Columns.AutoFit();

                System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheetSrc);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheetDst);
                xlWorkBookDst.Close(true);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkBookDst);
                xlApp.UserControl = true;
                xlApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject((object)xlApp);

                AddLog(Constant.ZULUANALYSE_STEP2_SUCCESS);
            }
            catch (Exception e)
            {
                AddLog(Constant.ZULUANALYSE_STEP2_FAILED);
                AddLog(e.Message);
                retval = false;
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            return retval;
        }

        private bool zuluAnalyzerStep3()
        {
            string dstFileName = "Downloads//All_履歴.xlsx";
            bool retVal = true;
            Excel.Application xlApp = null;
            try
            {
                xlApp = new Excel.Application();
                Excel.Workbook xlWorkBookDst = null;
                Excel.Worksheet worksheetDst = null;
                Excel.Worksheet worksheetSrc = null;
                xlApp.Visible = false;
                dstFileName = System.IO.Path.GetFullPath(dstFileName);

                xlWorkBookDst = xlApp.Workbooks.Open(dstFileName, 0, false, 5, "", "", false,
                    Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

                int nCount = xlWorkBookDst.Worksheets.Count;
                worksheetSrc = (Excel.Worksheet)xlWorkBookDst.Worksheets["③計算結果代入+DIV 0!項目削除, ⑤ロット数計算結果代入"];
                worksheetDst = (Excel.Worksheet)xlWorkBookDst.Worksheets.Add(defaultArg, worksheetSrc, defaultArg, defaultArg);
                worksheetDst.Name = "④年別計算資料";
                
                worksheetSrc.Activate();
                int currentRowSrc = worksheetSrc.Cells.Find("*", System.Reflection.Missing.Value,
                               System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                               Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                               false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

                //Pivotテーブルの作成
                Excel.Range dataRange = worksheetSrc.get_Range("A1", "R" + currentRowSrc);
                Excel.PivotCache cache = xlWorkBookDst.PivotCaches().Add(Excel.XlPivotTableSourceType.xlDatabase, dataRange);
                Excel.PivotTables tables = worksheetDst.PivotTables();
                Excel.PivotTable pt = tables.Add(cache, worksheetDst.Range["A1"], "Pivot", defaultArg, defaultArg);

                pt.Format(Excel.XlPivotFormatType.xlPTNone);
                pt.ShowValuesRow = false;

                // Row Fields
                Excel.PivotField fld = ((Excel.PivotField)pt.PivotFields("System ID"));
                fld.Orientation = Excel.XlPivotFieldOrientation.xlRowField;

                fld.LayoutForm = Excel.XlLayoutFormType.xlTabular;
                fld.LayoutCompactRow = true;
                fld.LayoutSubtotalLocation = Excel.XlSubtototalLocationType.xlAtTop;
                fld.set_Subtotals(1, false);
                fld = ((Excel.PivotField)pt.PivotFields("Currency"));
                fld.Orientation = Excel.XlPivotFieldOrientation.xlRowField;

                fld.LayoutForm = Excel.XlLayoutFormType.xlTabular;
                fld.LayoutCompactRow = true;
                fld.LayoutSubtotalLocation = Excel.XlSubtototalLocationType.xlAtTop;
                fld.set_Subtotals(1, false);
                fld = ((Excel.PivotField)pt.PivotFields("Type"));
                fld.Orientation = Excel.XlPivotFieldOrientation.xlRowField;
                fld.LayoutForm = Excel.XlLayoutFormType.xlOutline;
                fld.LayoutCompactRow = true;
                fld.LayoutSubtotalLocation = Excel.XlSubtototalLocationType.xlAtTop;
                fld.set_Subtotals(1, false);

                // Data Fields
                fld = ((Excel.PivotField)pt.PivotFields("Highest Profit(¥)"));
                fld.Orientation = Excel.XlPivotFieldOrientation.xlDataField;
                fld.Function = Excel.XlConsolidationFunction.xlSum;
                fld.NumberFormat = "¥#";
                fld = ((Excel.PivotField)pt.PivotFields("Worst Drawdown(¥)"));
                fld.Orientation = Excel.XlPivotFieldOrientation.xlDataField;
                fld.Function = Excel.XlConsolidationFunction.xlSum;
                fld.NumberFormat = "¥#";
                fld = (Excel.PivotField)pt.PivotFields("回数");
                fld.Orientation = Excel.XlPivotFieldOrientation.xlDataField;
                fld.Function = Excel.XlConsolidationFunction.xlSum;
                fld.Name = "回数計算";
                fld = (Excel.PivotField)pt.PivotFields("Profit(¥)");
                fld.Orientation = Excel.XlPivotFieldOrientation.xlDataField;
                fld.Function = Excel.XlConsolidationFunction.xlSum;
                fld.Name = "Profit(¥)計算";

                worksheetDst.Columns.AutoFit();

                System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheetSrc);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheetDst);
                xlWorkBookDst.Close(true);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkBookDst);
                xlApp.UserControl = true;
                xlApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject((object)xlApp);

                AddLog(Constant.ZULUANALYSE_STEP3_SUCCESS);
            }
            catch (Exception e)
            {
                AddLog(Constant.ZULUANALYSE_STEP3_FAILED);
                AddLog(e.Message);
                retVal = false;
            }
            GC.Collect();
            GC.WaitForPendingFinalizers();
            return retVal;

        }

        private bool zuluAnalyzerStep4()
        {
            bool retval = true;
            string dstFileName = "Downloads//All_履歴.xlsx";
            systems = new List<TradingSystem>();

            try
            {
                int currentRowSrc = 0;
                int currentColSrc = 0;
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkBookDst = null;
                Excel.Worksheet worksheetDst = null;
                Excel.Worksheet worksheetSrc = null;
                xlApp.Visible = false;
                dstFileName = System.IO.Path.GetFullPath(dstFileName);

                xlWorkBookDst = xlApp.Workbooks.Open(dstFileName, 0, false, 5, "", "", false,
                    Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

                worksheetSrc = (Excel.Worksheet)xlWorkBookDst.Worksheets["④年別計算資料"];
                worksheetDst = (Excel.Worksheet)xlWorkBookDst.Worksheets.Add(defaultArg, worksheetSrc, defaultArg, defaultArg);
                worksheetDst.Name = "⑤性能計算結果";

                currentRowSrc = worksheetSrc.Cells.Find("*", System.Reflection.Missing.Value,
                               System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                               Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                               false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row - 1;

                currentColSrc = worksheetSrc.Cells.Find("*", System.Reflection.Missing.Value,
                               System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                               Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                               false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Column;

                //データのコピー
                Excel.Range srcRng = worksheetSrc.get_Range("A2", "G" + currentRowSrc.ToString());
                Excel.Range dstRng = worksheetDst.get_Range("A2", "G" + currentRowSrc.ToString());
                srcRng.Copy();
                dstRng.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

                worksheetSrc.get_Range("D2", "D" + currentRowSrc.ToString()).NumberFormat = "¥#;[Red]-¥#";
                worksheetSrc.get_Range("E2", "E" + currentRowSrc.ToString()).NumberFormat = "¥#;[Red]-¥#";
                worksheetDst.get_Range("D2", "D" + currentRowSrc.ToString()).NumberFormat = "¥#;[Red]-¥#";
                worksheetDst.get_Range("E2", "E" + currentRowSrc.ToString()).NumberFormat = "¥#;[Red]-¥#";

                worksheetDst.Cells[1, 1].Value = worksheetSrc.Cells[1, 1].Value;
                worksheetDst.Cells[1, 2].Value = worksheetSrc.Cells[1, 2].Value;
                worksheetDst.Cells[1, 3].Value = worksheetSrc.Cells[1, 3].Value;
                worksheetDst.Cells[1, 4].Value = worksheetSrc.Cells[1, 4].Value;
                worksheetDst.Cells[1, 5].Value = worksheetSrc.Cells[1, 5].Value;
                worksheetDst.Cells[1, 6].value = "回数";
                worksheetDst.Cells[1, 7].value = "損益";
                worksheetDst.Cells[1, 8].value = "性能";

                Excel.Range r2 = worksheetDst.get_Range("H2", "H" + currentRowSrc.ToString());
                r2.Formula = "=IF(E2=0,0,D2/ABS(E2))";

                // Calculate the formula for whole worksheet
                worksheetDst.Calculate();
                worksheetDst.Activate();
                xlWorkBookDst.Save();

                AddLog("4段階の準備中です。");
               
                // Fill the blank cell and set filter.
                int currentRowDst = worksheetDst.Cells.Find("*", System.Reflection.Missing.Value,
                               System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                               Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                               false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

                // Fill the blank cell with above value in A colum
                Excel.Range firstEmpty = null;
                int i = 1;
                for (i = 2; i < currentRowDst; i++)
                {
                    firstEmpty = worksheetDst.get_Range("A" + i);
                    if (string.IsNullOrEmpty(firstEmpty.Value2))
                        break;
                }

                firstEmpty.Select();
                dstRng = worksheetDst.get_Range("A2", "A" + currentRowDst.ToString());
                Excel.Range target = dstRng.SpecialCells(Excel.XlCellType.xlCellTypeBlanks, Type.Missing);
                if (target != null)
                {
                    firstEmpty.Select();
                    target.Formula = string.Format("=A{0}", i - 1);
                    worksheetDst.Calculate();
                    xlWorkBookDst.Save();
                }

                // Fill the blank cell with above value in B colum
                firstEmpty = null;
                i = 1;
                for (i = 2; i < currentRowDst; i++)
                {
                    firstEmpty = worksheetDst.get_Range("B" + i);
                    if (string.IsNullOrEmpty(firstEmpty.Value2))
                        break;
                }

                firstEmpty.Select();
                dstRng = worksheetDst.get_Range("B2", "B" + currentRowDst.ToString());
                target = dstRng.SpecialCells(Excel.XlCellType.xlCellTypeBlanks, Type.Missing);
                if (target != null)
                {
                    firstEmpty.Select();
                    target.Formula = string.Format("=B{0}", i - 1);
                    worksheetDst.Calculate();
                    xlWorkBookDst.Save();
                }

                worksheetDst.Columns.AutoFit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheetSrc);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheetDst);
                xlWorkBookDst.Close(true);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkBookDst);
                xlApp.UserControl = true;
                xlApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject((object)xlApp);

                AddLog(Constant.ZULUANALYSE_STEP4_SUCCESS);
            }
            catch (Exception e)
            {
                AddLog(Constant.ZULUANALYSE_STEP4_FAILED);
                AddLog(e.Message);
                retval = false;
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            return retval;
        }

        /*
         * Final_履歴ファイルを作成
         */
        private bool zuluAnalyzerStep5()
        {
            bool retval = true;
            string dstFileName = "Downloads//Final_履歴.xlsx";
            string srcFileName = "Downloads//ALL_履歴.xlsx";
            systems = new List<TradingSystem>();

            try
            {
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkBookDst = null;
                Excel.Worksheet worksheetDst = null;
                Excel.Workbook xlWorkBookSrc = null;
                Excel.Worksheet worksheetSrc = null;
                xlApp.Visible = false;
                dstFileName = System.IO.Path.GetFullPath(dstFileName);
                srcFileName = System.IO.Path.GetFullPath(srcFileName);

                if (File.Exists(dstFileName))
                    File.Delete(dstFileName);

                xlWorkBookDst = xlApp.Workbooks.Add(Type.Missing);
                xlWorkBookSrc = xlApp.Workbooks.Open(srcFileName, 0, false, 5, "", "", false,
                    Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                worksheetSrc = (Excel.Worksheet)xlWorkBookSrc.ActiveSheet;
                worksheetDst = (Excel.Worksheet)xlWorkBookDst.ActiveSheet;
                worksheetDst.Name = "性能計算結果";

                int currentRowSrc = worksheetSrc.Cells.Find("*", System.Reflection.Missing.Value,
                               System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                               Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                               false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

                /*
                 * データのコピー
                 */
                Excel.Range srcRng = worksheetSrc.get_Range("A1", "H" + currentRowSrc.ToString());
                Excel.Range dstRng = worksheetDst.get_Range("A1", Type.Missing);
                srcRng.Copy();
                dstRng.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                xlWorkBookDst.SaveAs(dstFileName);

                int currentRowDst = worksheetDst.Cells.Find("*", System.Reflection.Missing.Value,
                               System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                               Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                               false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

                // Sort by profit as Desc
                dynamic allDataRange = worksheetDst.get_Range("A2", "H" + currentRowDst.ToString());
                allDataRange.Sort(allDataRange.Columns[8], Excel.XlSortOrder.xlDescending);
                xlWorkBookDst.Save();
                
                // Set Filters
                allDataRange = worksheetDst.UsedRange;
                allDataRange.AutoFilter(1, "<>", Excel.XlAutoFilterOperator.xlFilterValues, Type.Missing, true);
                Excel.Range filteredRange = allDataRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible, Excel.XlSpecialCellsValue.xlTextValues);
                xlWorkBookDst.Save();

                string trader, currency, type, high, worst, count, profit, spec;
                // Update the listTrading
                for (int i = 2; i <= currentRowDst; i++)
                {
                    trader = worksheetDst.Cells[i, 1].Value2.ToString();
                    currency = worksheetDst.Cells[i, 2].Value2.ToString();
                    type = worksheetDst.Cells[i, 3].Value2.ToString();
                    high = worksheetDst.Cells[i, 4].Value2.ToString();
                    worst = worksheetDst.Cells[i, 5].Value2.ToString();
                    count = worksheetDst.Cells[i, 6].Value2.ToString();
                    profit = worksheetDst.Cells[i, 7].Value2.ToString();
                    spec = worksheetDst.Cells[i, 8].Value2.ToString();

                    /*
                     * 性能が良い５００個のシステムだけ追加
                     */
                    if (i > 500)
                        break;
                    
                    TradingSystem tSystem = new TradingSystem();
                    tSystem.trader = trader;
                    tSystem.currency = currency;
                    tSystem.type = type;
                    tSystem.high = float.Parse(high);
                    tSystem.worst = float.Parse(worst);
                    tSystem.count = int.Parse(count);
                    tSystem.profit = float.Parse(profit);
                    tSystem.spec = float.Parse(spec);
                    tSystem.lot = 0.01f;
                    tSystem.maxLot = 0;
                    tSystem.maxWorst = 0;
                    systems.Add(tSystem);
                    
                    addTradingListView(trader, currency, type, high, worst, count, profit, spec, tSystem.maxLot, tSystem.maxWorst);
                }

                worksheetDst.Columns.AutoFit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheetDst);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheetSrc);
                xlWorkBookDst.Close(true);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkBookDst);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkBookSrc);
                xlApp.UserControl = true;
                xlApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject((object)xlApp);

                AddLog(Constant.ZULUANALYSE_STEP5_SUCCESS);
            }
            catch (Exception e)
            {
                AddLog(Constant.ZULUANALYSE_STEP5_FAILED);
                AddLog(e.Message);
                retval = false;
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            return retval;
        }

        public void resultInCallback(bool result)
        {
            if (result)
                MessageBox.Show(Constant.ERROR_MESSAGE_SUCCESS_READ);
            else
                MessageBox.Show(Constant.ERROR_MESSAGE_INVALID_FILE);
        }

        /*
         * 相場データをダウンロードする
         */
        private bool downloadExchangeRate()
        {
            bool ret = true;
            http_request.setURL("https://financialmodelingprep.com/api/v3/forex");
            http_request.setSendMode(HTTP_SEND_MODE.HTTP_GET);
            http_request.setUserAgent("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/78.0.3904.97 Safari/537.36");
            http_request.appendCustomHeader("Accept-Language: ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7");
            http_request.appendCustomHeader("Accept-Encoding: gzip, deflate, br");
            http_request.appendCustomHeader("Cache-Control: max-age=0");
            http_request.appendCustomHeader("Upgrade-Insecure-Requests:	1");
            http_request.appendCustomHeader("Sec-Fetch-Site: cross-site");
            http_request.appendCustomHeader("Sec-Fetch-Mode: navigate");
            http_request.setAccept("text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3");
            http_request.setAutoRedirect(true);

            if (!http_request.sendRequest(false, "", "application/json;charset=UTF-8"))
            {
                AddLog(Constant.ZULUANALYSE_STEP0_FAILED);
                ret = false;
            }
            else
            {
                try
                {
                    currentList = new List<string>();
                    
                    string response = http_request.getResponseString();
                    exchageRate = JsonConvert.DeserializeObject<ExchangeClass.ExchangeRate>(response);

                    AddLog(Constant.ZULUANALYSE_STEP0_SUCCESS);
                    AddLog(string.Format(Constant.ZULUANALYSE_EXCHANGE_RATE, exchageRate.forexList[0].date));

                    /*
                     * JPY 通貨の相場データしりょう
                     */
                    foreach (ExchangeClass.ExchangeTable table in exchageRate.forexList)
                    {
                        if (table.ticker.Contains("JPY"))
                        {
                            if (!currentList.Contains(table.ticker))
                                currentList.Add(table.ticker);

                            AddLog("" + table.ticker + ":" + table.open);
                        }
                    }
                }
                catch (Exception e)
                {
                    AddLog(Constant.ZULUANALYSE_STEP0_FAILED);
                    ret = false;
                }
                
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();

            return ret;
        }

        private void btnCombine_Click(object sender, EventArgs e)
        {
            
        }

        private void analyzeCreateAllHistory()
        {
            Invoke((MethodInvoker)(() => processCombineAction()));
        }

        private void processCombineAction()
        {
            /*
             * 相場データをがない場合は、相場データをダウンロードする。
             */
            if (currentList == null || currentList.Count == 0)
            {
                if (!downloadExchangeRate())
                    return;
            }

            /*
             * ELSXファイルの最初の列にトレードの名前を入れる
             */
            Invoke((MethodInvoker)(() => progressBar1.Value = 0));
            if (!contentShiftRightExcel())
                return;

            Invoke((MethodInvoker)(() => progressBar1.Value = currentDownload));
            

            string dstFileName = "Downloads//All_履歴.xlsx";
            
            if (File.Exists(dstFileName))
                File.Delete(dstFileName);

            /*
             * トレードのデータを結合。
             * メモリ消費を減らすため、50個ずつ読む。
             */
            for (int i = 0; i < fileListView.Items.Count; i += 50)
            {
                if (!mergeAllHistory(i))
                    return;
            }
            progressBar1.Value += 1;

            if (!deleteColumn())
                return;

            /*
             * Exlsシートの「②計算資料」を計算
             */
            Invoke((MethodInvoker)(() => progressBar1.Value = currentDownload + 1));
            if (!zuluAnalyzerStep1())
                return;

            /*
             * Exlsシートの「③計算結果代入+DIV 0!項目削除, ⑤ロット数計算結果代入」を計算
             */
            Invoke((MethodInvoker)(() => progressBar1.Value = currentDownload + 2));
            if (!zuluAnalyzerStep2())
                return;

            /*
             * Exlsシートの「④年別計算資料」を計算
             */
            Invoke((MethodInvoker)(() => progressBar1.Value = currentDownload + 3));
            if (!zuluAnalyzerStep3())
                return;

            /*
             * Exlsシートの「⑤性能計算結果」を計算
             */
            Invoke((MethodInvoker)(() => progressBar1.Value = currentDownload + 4));
            if (!zuluAnalyzerStep4())
                return;

            if (!zuluAnalyzerStep5())
                return;

            Invoke((MethodInvoker)(() => progressBar1.Value = currentDownload + 5));

            dstFileName = "Downloads//Final_履歴.xlsx";
            dstFileName = System.IO.Path.GetFullPath(dstFileName);
            System.Diagnostics.Process.Start(dstFileName);
            btnCombine.Enabled = true;
        }

        private void fileListView_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        private void listFileNames(object sender, EventArgs e)
        {
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            zuluAnalyzerStep4();
        }

        private void tpgDownload_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                int currentRowDst = 0;
                int currentColDst = 0;
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkBookDst = null;
                Excel.Worksheet worksheetDst = null;
                xlApp.Visible = false;

                string dstFileName = "Downloads//2.xlsx";

                dstFileName = System.IO.Path.GetFullPath(dstFileName);
                xlWorkBookDst = xlApp.Workbooks.Open(dstFileName, 0, false, 5, "", "", false,
                    Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                worksheetDst = (Excel.Worksheet)xlWorkBookDst.ActiveSheet;

                currentRowDst = worksheetDst.Cells.Find("*", System.Reflection.Missing.Value,
                                System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                                Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                                false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

                currentColDst = worksheetDst.Cells.Find("*", System.Reflection.Missing.Value,
                                System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                                Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                                false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Column;

                int limit = 1000000;
                int sheetColCount = worksheetDst.Cells.Find("*", System.Reflection.Missing.Value,
                                System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                                Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlPrevious,
                                false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Column;

                Excel.Range range1 = worksheetDst.get_Range("P2", "P" + currentRowDst.ToString());
                range1.Cells.NumberFormat = "$#";

                System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheetDst);
                xlWorkBookDst.Close(true);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkBookDst);
                xlApp.UserControl = true;
                xlApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject((object)xlApp);

                progressBar1.Value += 1;
            }
            catch (Exception ex)
            {
                AddLog(ex.Message);
            }
            finally
            {

            }
        }
        
        private void btnCombine_Click_1(object sender, EventArgs e)
        {
            progressBar1.Maximum = currentDownload + 5;
            progressBar1.Step = 1;
            progressBar1.Value = 0;

            Thread thread = new Thread(new ThreadStart(analyzeCreateAllHistory));
            thread.Start();

            btnCombine.Enabled = false;
        }
        /*
         *性能が計算され、ランキングされたファイルの読み取り
         */
        private void readFromCalculatedFile(object sender, EventArgs e)
        {
            string dstFileName = "Downloads//Final_履歴.xlsx";

            try
            {
                dstFileName = System.IO.Path.GetFullPath(dstFileName);
                if (!File.Exists(dstFileName))
                {
                    MessageBox.Show(Constant.ERROR_MESSAGE_NO_FILE);
                    return;
                }
                AddLog("計算した履歴ファイルからデータを読んでいます。");

                callbackFormReadHistoryFile callback = new callbackFormReadHistoryFile(resultInCallback);
                Thread readThread = new Thread(() => readHistoryFile(callback));
                readThread.Start();

            }
            catch (Exception ex)
            {
                MessageBox.Show(Constant.ERROR_MESSAGE_FAILED_READ);
                return;
            }
        }

        private void readHistoryFile(callbackFormReadHistoryFile callback)
        {
            callbackFormReadHistoryFile _callback = callback;
            bool retval = true;
            string dstFileName = "Downloads//Final_履歴.xlsx";
            systems = new List<TradingSystem>();

            int currentRowSrc = 0;
            int currentColSrc = 0;
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkBookDst = null;
            Excel.Worksheet worksheetDst = null;
            xlApp.Visible = false;
            dstFileName = System.IO.Path.GetFullPath(dstFileName);

            try
            {
                xlWorkBookDst = xlApp.Workbooks.Open(dstFileName, 0, false, 5, "", "", false,
                    Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

                worksheetDst = (Excel.Worksheet)xlWorkBookDst.Worksheets["性能計算結果"];

                currentRowSrc = worksheetDst.Cells.Find("*", System.Reflection.Missing.Value,
                                System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                                Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                                false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row - 1;

                currentColSrc = worksheetDst.Cells.Find("*", System.Reflection.Missing.Value,
                                System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                                Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                                false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Column;

                Excel.Range dstRng = worksheetDst.get_Range("A2", "G" + currentRowSrc.ToString());
                dstRng.Copy();
                dstRng.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

                string trader, currency, type, high, worst, count, profit, spec;
                // Update the listTrading

                /*
                 * 計算に全部参加しないので、性能が良い５００個のシステムだけ追加する
                 */
                for (int i = 2; i <= currentRowSrc; i++)
                {
                    if (i > 500)
                        break;

                    trader = worksheetDst.Cells[i, 1].Value2.ToString();
                    currency = worksheetDst.Cells[i, 2].Value2.ToString();
                    type = worksheetDst.Cells[i, 3].Value2.ToString();
                    high = worksheetDst.Cells[i, 4].Value2.ToString();
                    worst = worksheetDst.Cells[i, 5].Value2.ToString();
                    count = worksheetDst.Cells[i, 6].Value2.ToString();
                    profit = worksheetDst.Cells[i, 7].Value2.ToString();

                    if (profit.StartsWith("-"))
                        continue;

                    spec = worksheetDst.Cells[i, 8].Value2.ToString();

                    TradingSystem tSystem = new TradingSystem();
                    tSystem.trader = trader;
                    tSystem.currency = currency;
                    tSystem.type = type;
                    tSystem.high = float.Parse(high);
                    tSystem.worst = float.Parse(worst);
                    tSystem.count = int.Parse(count);
                    tSystem.profit = float.Parse(profit);
                    tSystem.spec = float.Parse(spec);
                    tSystem.lot = 0.01f;
                    tSystem.maxLot = 0;
                    tSystem.maxWorst = 0;
                    systems.Add(tSystem);

                    xlWorkBookDst.Application.CutCopyMode = 0;

                    addTradingListView(trader, currency, type, high, worst, count, profit, spec, tSystem.maxLot, tSystem.maxWorst);
                }
                AddLog(Constant.ZULUANALYSE_STEP4_SUCCESS);
            }
            catch (Exception ex)
            {
                AddLog(Constant.ZULUANALYSE_STEP4_FAILED);
                AddLog(ex.Message);
                retval = false;
            }
            finally
            {
                worksheetDst.Columns.AutoFit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheetDst);
                xlWorkBookDst.Close(true);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkBookDst);
                xlApp.UserControl = true;
                xlApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject((object)xlApp);
            }

            _callback(retval);
            return;
        }

        private async void btnFinal_Click(object sender, EventArgs e)
        {
            if (Convert.ToInt32(comboBox1.SelectedValue) == -1)
            {
                MessageBox.Show(Constant.ERROR_MESSAGE_SELECT_COUNT);
                return;
            }

            Thread thread = new Thread(() => kumiawaseWork());
            thread.Start();
        }

        


        private void kumiawaseWork()
        {
            Invoke((MethodInvoker)(() => kumiawase()));
        }

        private void kumiawase()
        {
            listRanking.Items.Clear();

            int minTradingCount = int.Parse(textMinTraderCount.Text);
            int minProfit = int.Parse(textProfit.Text);
            int maxDrawDown = int.Parse(textMaxDrawdown.Text);
            int systemCount = Convert.ToInt32(comboBox1.SelectedValue);

            resortSystems(minTradingCount);

            LinkedList<KumiawaseItem> kumiawases = new LinkedList<KumiawaseItem>();
            KumiawaseItem item = new KumiawaseItem();

            /*
             * 取引数が条件に合わないシステムは除去する。
             */
            foreach (ListViewItem lvi in listTrading.Items)
            {
                int value = int.Parse(lvi.SubItems[5].Text);
                if (minTradingCount > value)
                    lvi.Checked = false;
                lvi.BackColor = default(Color);
            }


            /*
             * 方程式の桁数を保管する変数
             */
            float[] a = { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
            float[] b = { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
            int[] xup = { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
            Array.Clear(c, 0, 20);
            
            int i = 0;
            bool skipflag = false;
            int skipCount = 0;

            foreach (ListViewItem lvi in listTrading.Items)
            {
                int index = lvi.Index;
                if (lvi.Checked == false)
                    continue;

                int halfCount = systemCount / 2;
                if (skipflag)
                {
                    skipCount++;
                    if (skipCount == halfCount)
                    {
                        skipflag = false;
                    }
                    continue;
                }

                /*
                 * 取引一度の損益、装荷を計算する。
                 */
                a[i] = systems[index].profit / systems[index].count;
                b[i] = Math.Abs(systems[index].worst / systems[index].count);
                c[i] = index;

                double lot = 0;
                bool chk = Double.TryParse(lvi.SubItems[9].Text, out lot);
                if (!chk)
                {
                    return;
                }

                double w = 0;
                chk = Double.TryParse(lvi.SubItems[10].Text, out w);
                if (!chk)
                {
                    return;
                }

                systems[i].maxLot = (float)lot;
                systems[i].maxWorst = (float)w;

                i++;


                if (i == halfCount)
                    skipflag = true;

                if (i == (systemCount))
                    break;
            }

            int xupdasi = (int)(maxDrawDown / (b[0] + b[1] + b[2] + b[3] + b[4] + b[5] + b[6] + b[7] + b[8] + b[9] + b[10] + b[11] + b[12] + b[13] + b[14] + b[15] + b[16] + b[17] + b[18] + b[19]));

            int width = 10;
            switch (systemCount)
            {
                case 3:
                    width = 200;
                    break;
                case 5://9,765,625
                    width = 250;
                    break;
                case 7://10,000,000
                    width = 150;
                    break;
                case 9://10,077,696
                    width = 150;
                    break;
                case 10://9,765,625
                    width = 120;
                    break;
                case 15:
                    width = 100;
                    break;
                case 20:
                    width = 80;
                    break;

            }
            int xupstart = xupdasi - width;

            /*
             * システムの開始ロット値を決定
             */
            for (i = 0; i < systemCount; i ++)
            {
                xup[i] = (int)((maxDrawDown + b[i] - b[0] - b[1] - b[2] - b[3] - b[4]
                    - b[5] - b[6] - b[7] - b[8] - b[9] -b[10] - b[11] - b[12] - b[13]
                    - b[14] - b[15] - b[16] - b[17] - b[18] - b[19]) / b[i]);

                if (systems[i].maxLot != 0)
                    xup[i] = xup[i] < (int)(systems[i].maxLot / 0.01) ? xup[i] : (int)(systems[i].maxLot / 0.01);

                if (systems[i].maxWorst != 0)
                    xup[i] = xup[i] < Math.Abs((int)(systems[i].maxWorst / systems[i].worst)) ? xup[i] : Math.Abs((int)(systems[i].maxWorst / systems[i].worst));

                xup[i] = xup[i] < xupdasi ? xup[i] : xupdasi;
                if (xup[i] > 100)
                    xup[i] -= i;
            }

            int[] j = { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
            int[] jtemp = { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
            bool flag = true;
            float sum = 0;
            float worst = 0;
            float sp = 0;
            int s = 0;

            float maxSpec = 0;
            float tmpSpec = 0;
            int[] optimizeJ = { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };

            for (i = 0; i < systemCount; i++)
            {
                j[i] = xupstart - i;
                if (j[i] <= 0) j[i] = 1;
            }
                
            do
            {
                j[19]++;
                do
                {
                    j[18]++;
                    do
                    {
                        j[17]++;
                        do
                        {
                            j[16]++;
                            do
                            {
                                j[15]++;
                                do
                                {
                                    j[14]++;
                                    do
                                    {
                                        j[13]++;
                                        do
                                        {
                                            j[12]++;
                                            do
                                            {
                                                j[11]++;
                                                do
                                                {
                                                    j[10]++;
                                                    do
                                                    {
                                                        j[9]++;
                                                        do
                                                        {
                                                            j[8]++;
                                                            do
                                                            {
                                                                j[7]++;
                                                                do
                                                                {
                                                                    j[6]++;
                                                                    do
                                                                    {
                                                                        j[5]++;
                                                                        do
                                                                        {
                                                                            j[4]++;
                                                                            do
                                                                            {
                                                                                j[3]++;
                                                                                do
                                                                                {
                                                                                    j[2]++;
                                                                                    do
                                                                                    {
                                                                                        j[1]++;
                                                                                        do
                                                                                        {
                                                                                            j[0]++;
                                                                                            sum = 0;
                                                                                            for (s = 0; s < 20; s++)
                                                                                                sum += (a[s] * j[s]);
                                                                                            if (sum > minProfit)
                                                                                            {
                                                                                                worst = 0;
                                                                                                for (s = 0; s < 20; s++)
                                                                                                    worst += (b[s] * j[s]);

                                                                                                if (worst < maxDrawDown)
                                                                                                {
                                                                                                    //flag = false;
                                                                                                    //break;

                                                                                                    tmpSpec = 0;
                                                                                                    tmpSpec = systems[0].spec * j[0] + systems[1].spec * j[1] + systems[2].spec * j[2] + systems[3].spec * j[3] + systems[4].spec * j[4] +
                                                                                                        systems[5].spec * j[5] + systems[6].spec * j[6] + systems[7].spec * j[7] + systems[8].spec * j[8] + systems[9].spec * j[9] +
                                                                                                        systems[10].spec * j[10] + systems[11].spec * j[11] + systems[12].spec * j[12] + systems[13].spec * j[13] + systems[14].spec * j[14] +
                                                                                                        systems[15].spec * j[15] + systems[16].spec * j[16] + systems[17].spec * j[17] + systems[18].spec * j[18] + systems[19].spec * j[19];

                                                                                                    if (tmpSpec > maxSpec)
                                                                                                    {
                                                                                                        optimizeJ[0] = j[0];
                                                                                                        optimizeJ[1] = j[1];
                                                                                                        optimizeJ[2] = j[2];
                                                                                                        optimizeJ[3] = j[3];
                                                                                                        optimizeJ[4] = j[4];
                                                                                                        optimizeJ[5] = j[5];
                                                                                                        optimizeJ[6] = j[6];
                                                                                                        optimizeJ[7] = j[7];
                                                                                                        optimizeJ[8] = j[8];
                                                                                                        optimizeJ[9] = j[9];
                                                                                                        optimizeJ[10] = j[10];
                                                                                                        optimizeJ[11] = j[11];
                                                                                                        optimizeJ[12] = j[12];
                                                                                                        optimizeJ[13] = j[13];
                                                                                                        optimizeJ[14] = j[14];
                                                                                                        optimizeJ[15] = j[15];
                                                                                                        optimizeJ[16] = j[16];
                                                                                                        optimizeJ[17] = j[17];
                                                                                                        optimizeJ[18] = j[18];
                                                                                                        optimizeJ[19] = j[19];

                                                                                                        sp = 0;
                                                                                                        for (int t = 0; t < 20; t++) {
                                                                                                            sp += (systems[t].spec * j[t]);
                                                                                                        }
                                                                                                            
                                                                                                        item = new KumiawaseItem();
                                                                                                        item.spec = sp;

                                                                                                        for (int t = 0; t < 20; t++){
                                                                                                            jtemp[t] = j[t];
                                                                                                            if (jtemp[t] > 999)
                                                                                                                jtemp[t] = jtemp[t] % 1000;
                                                                                                        }

                                                                                                        /*
                                                                                                         * 計算結果のセーブ
                                                                                                         * こちらはエラーです。
                                                                                                         */
                                                                                                        /////////////////////////////////////////////////////////


                                                                                                        //////////////////////////////////////////////////////////
                                                                                                        kumiawases.AddFirst(item);
                                                                                                        if (kumiawases.Count == 10)
                                                                                                            kumiawases.RemoveLast();
                                                                                                    }

                                                                                                }
                                                                                            }

                                                                                        } while ((j[0] < xup[0]) && flag);

                                                                                    } while ((j[1] < xup[1]) && flag);

                                                                                } while ((j[2] < xup[2]) && flag);

                                                                            } while ((j[3] < xup[3]) && flag);

                                                                        } while ((j[4] < xup[4]) && flag);

                                                                    } while ((j[5] < xup[5]) && flag);

                                                                } while ((j[6] < xup[6]) && flag);

                                                            } while ((j[7] < xup[7]) && flag);

                                                        } while ((j[8] < xup[8]) && flag);

                                                    } while ((j[9] < xup[9]) && flag);

                                                } while ((j[10] < xup[10]) && flag);

                                            } while ((j[11] < xup[11]) && flag);

                                        } while ((j[12] < xup[12]) && flag);

                                    } while ((j[13] < xup[13]) && flag);

                                } while ((j[14] < xup[14]) && flag);

                            } while ((j[15] < xup[15]) && flag);

                        } while ((j[16] < xup[16]) && flag);

                    } while ((j[17] < xup[17]) && flag);

                } while ((j[18] < xup[18]) && flag);

            } while ((j[19] < xup[19]) && flag);

            linearList = new List<KumiawaseItem>();

            if (kumiawases.Count == 0)
            {
                MessageBox.Show(Constant.ERROR_MESSAGE_NOT_AWASE);
                return;
            }
            for (i = 0; i < kumiawases.Count; i++)
                linearList.Add(kumiawases.ElementAt(i));


            /*
             * システムのソート
             */
            linearList.Sort(delegate (KumiawaseItem item1, KumiawaseItem item2) { return item1.spec >= item2.spec ? -1 : 1; });

            for (i = 0; i < kumiawases.Count; i ++)
            {
                ListViewItem lvi = new ListViewItem();
                lvi.Text = i.ToString();
                lvi.SubItems.Add(string.Format("グループ{0}", i));
                lvi.SubItems.Add(linearList[i].spec.ToString());
                listRanking.Items.Add(lvi);
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
        private void kumiawaseWork2()
        {
            Invoke((MethodInvoker)(() => kumiawase2()));
        }

private void kumiawase2()
        {
            listRanking2.Items.Clear();

            if (listRanking.Items.Count == 0)
            {
                MessageBox.Show("「結合１」のボタンを押してください。");
                return;
            }

            int minTradingCount = int.Parse(textMinTraderCount.Text);
            int minProfit = int.Parse(textProfit2.Text);
            int maxDrawDown = int.Parse(textMaxDrawDown2.Text);
            int systemCount = Convert.ToInt32(comboBox1.SelectedValue);

            LinkedList<KumiawaseItem> kumiawases = new LinkedList<KumiawaseItem>();
            KumiawaseItem item = new KumiawaseItem();

            /*
             * 取引数が条件に合わないシステムは除去する。
             */
            foreach (ListViewItem lvi in listTrading.Items)
            {
                int value = int.Parse(lvi.SubItems[5].Text);
                if (minTradingCount > value)
                    lvi.Checked = false;
                lvi.BackColor = default(Color);
            }

            /*
             * 方程式の桁数を保管する変数
             */
            float[] a = { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
            float[] b = { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
            int[] xup = { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
            Array.Clear(c2, 0, 20);

            int i = 0;
            bool skipflag = true;
            int skipCount = 0;

            foreach (ListViewItem lvi in listTrading.Items)
            {
                int index = lvi.Index;
                if (lvi.Checked == false)
                    continue;

                /*
                 * システムを半分ずつ二つの部分に分けて計算する。
                 */
                int halfCount = systemCount / 2;
                if (skipflag)
                {
                    skipCount++;
                    if (skipCount == halfCount)
                    {
                        skipflag = false;
                    }

                    if (skipCount == systemCount)
                    {
                        skipflag = false;
                    }
                    continue;
                }

                /*
                 * 取引一度の損益、装荷を計算する。
                 */
                a[i] = systems[index].profit / systems[index].count;
                b[i] = Math.Abs(systems[index].worst / systems[index].count);
                c2[i] = index;

                double lot = 0;
                bool chk = Double.TryParse(lvi.SubItems[9].Text, out lot);
                if (!chk)
                {
                    return;
                }

                double w = 0;
                chk = Double.TryParse(lvi.SubItems[10].Text, out w);
                if (!chk)
                {
                    return;
                }

                systems[i].maxLot = (float)lot;
                systems[i].maxWorst = (float)w;

                i++;

                if (index == systemCount - 1)
                    skipflag = true;

                if (i == (systemCount))
                    break;
            }

            /*
             * 全システムが一緒に参加すると仮定し平均ロット数を求める。
             */
            int xupdasi = (int)(maxDrawDown / (b[0] + b[1] + b[2] + b[3] + b[4] + b[5] + b[6] + b[7] + b[8] + b[9] + b[10] + b[11] + b[12] + b[13] + b[14] + b[15] + b[16] + b[17] + b[18] + b[19]));

            int width = 10;

            /*
             * ロットの変化幅を指定する。
             */
            switch (systemCount)
            {
                case 3:
                    width = 200;
                    break;
                case 5://9,765,625
                    width = 250;
                    break;
                case 7://10,000,000
                    width = 150;
                    break;
                case 9://10,077,696
                    width = 150;
                    break;
                case 10://9,765,625
                    width = 120;
                    break;
                case 15:
                    width = 100;
                    break;
                case 20:
                    width = 80;
                    break;

            }
            int xupstart = xupdasi - width;

            /*
             * システムの開始ロット値を決定
             */
            for (i = 0; i < systemCount; i++)
            {
                xup[i] = (int)((maxDrawDown + b[i] - b[0] - b[1] - b[2] - b[3] - b[4]
                    - b[5] - b[6] - b[7] - b[8] - b[9] - b[10] - b[11] - b[12] - b[13]
                    - b[14] - b[15] - b[16] - b[17] - b[18] - b[19]) / b[i]);

                if (systems[i].maxLot != 0)
                    xup[i] = xup[i] < (int)(systems[i].maxLot / 0.01) ? xup[i] : (int)(systems[i].maxLot / 0.01);

                if (systems[i].maxWorst != 0)
                    xup[i] = xup[i] < Math.Abs((int)(systems[i].maxWorst / systems[i].worst)) ? xup[i] : Math.Abs((int)(systems[i].maxWorst / systems[i].worst));

                /*
                 * 開始ロット値が下限を超えないようにする
                 */
                xup[i] = xup[i] < xupdasi ? xup[i] : xupdasi;
            }

            int[] j = { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
            bool flag = true;
            float sum = 0;
            float worst = 0;
            float sp = 0;
            int s = 0;

            float maxSpec = 0;
            float tmpSpec = 0;
            int[] optimizeJ = { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };


            do
            {
                j[19]++;
                do
                {
                    j[18]++;
                    do
                    {
                        j[17]++;
                        do
                        {
                            j[16]++;
                            do
                            {
                                j[15]++;
                                do
                                {
                                    j[14]++;
                                    do
                                    {
                                        j[13]++;
                                        do
                                        {
                                            j[12]++;
                                            do
                                            {
                                                j[11]++;
                                                do
                                                {
                                                    j[10]++;
                                                    do
                                                    {
                                                        j[9]++;
                                                        do
                                                        {
                                                            j[8]++;
                                                            do
                                                            {
                                                                j[7]++;
                                                                do
                                                                {
                                                                    j[6]++;
                                                                    do
                                                                    {
                                                                        j[5]++;
                                                                        do
                                                                        {
                                                                            j[4]++;
                                                                            do
                                                                            {
                                                                                j[3]++;
                                                                                do
                                                                                {
                                                                                    j[2]++;
                                                                                    do
                                                                                    {
                                                                                        j[1]++;
                                                                                        do
                                                                                        {
                                                                                            j[0]++;
                                                                                            sum = 0;

                                                                                            /*
                                                                                             * システム結合の損益を計算
                                                                                             */
                                                                                            for (s = 0; s < 20; s++)
                                                                                                sum += (a[s] * j[s]);
                                                                                            if (sum > minProfit)
                                                                                            {
                                                                                                worst = 0;

                                                                                                /*
                                                                                                 * システム結合の装荷を計算
                                                                                                 */
                                                                                                for (s = 0; s < 20; s++)
                                                                                                    worst += (b[s] * j[s]);

                                                                                                if (worst < maxDrawDown)
                                                                                                {
                                                                                                    //flag = false;
                                                                                                    //break;

                                                                                                    tmpSpec = 0;

                                                                                                    /*
                                                                                                     * システム結合の性能を計算
                                                                                                     */
                                                                                                    tmpSpec = systems[0].spec * j[0] + systems[1].spec * j[1] + systems[2].spec * j[2] + systems[3].spec * j[3] + systems[4].spec * j[4] +
                                                                                                        systems[5].spec * j[5] + systems[6].spec * j[6] + systems[7].spec * j[7] + systems[8].spec * j[8] + systems[9].spec * j[9] +
                                                                                                        systems[10].spec * j[10] + systems[11].spec * j[11] + systems[12].spec * j[12] + systems[13].spec * j[13] + systems[14].spec * j[14] +
                                                                                                        systems[15].spec * j[15] + systems[16].spec * j[16] + systems[17].spec * j[17] + systems[18].spec * j[18] + systems[19].spec * j[19];

                                                                                                    if (tmpSpec > maxSpec)
                                                                                                    {
                                                                                                        optimizeJ[0] = j[0];
                                                                                                        optimizeJ[1] = j[1];
                                                                                                        optimizeJ[2] = j[2];
                                                                                                        optimizeJ[3] = j[3];
                                                                                                        optimizeJ[4] = j[4];
                                                                                                        optimizeJ[5] = j[5];
                                                                                                        optimizeJ[6] = j[6];
                                                                                                        optimizeJ[7] = j[7];
                                                                                                        optimizeJ[8] = j[8];
                                                                                                        optimizeJ[9] = j[9];
                                                                                                        optimizeJ[10] = j[10];
                                                                                                        optimizeJ[11] = j[11];
                                                                                                        optimizeJ[12] = j[12];
                                                                                                        optimizeJ[13] = j[13];
                                                                                                        optimizeJ[14] = j[14];
                                                                                                        optimizeJ[15] = j[15];
                                                                                                        optimizeJ[16] = j[16];
                                                                                                        optimizeJ[17] = j[17];
                                                                                                        optimizeJ[18] = j[18];
                                                                                                        optimizeJ[19] = j[19];

                                                                                                        sp = 0;
                                                                                                        for (int t = 0; t < 20; t++)
                                                                                                            sp += (systems[t].spec * j[t]);
                                                                                                        item = new KumiawaseItem();
                                                                                                        item.spec = sp;

                                                                                                        /*
                                                                                                         * 計算結果のセーブ
                                                                                                         * こちらはエラーです。
                                                                                                         */
                                                                                                        /////////////////////////////////////////////////////////


                                                                                                        //////////////////////////////////////////////////////////
                                                                                                        kumiawases.AddFirst(item);
                                                                                                        if (kumiawases.Count == 10)
                                                                                                            kumiawases.RemoveLast();
                                                                                                    }

                                                                                                }
                                                                                            }

                                                                                        } while ((j[0] < xup[0]) && flag);

                                                                                    } while ((j[1] < xup[1]) && flag);

                                                                                } while ((j[2] < xup[2]) && flag);

                                                                            } while ((j[3] < xup[3]) && flag);

                                                                        } while ((j[4] < xup[4]) && flag);

                                                                    } while ((j[5] < xup[5]) && flag);

                                                                } while ((j[6] < xup[6]) && flag);

                                                            } while ((j[7] < xup[7]) && flag);

                                                        } while ((j[8] < xup[8]) && flag);

                                                    } while ((j[9] < xup[9]) && flag);

                                                } while ((j[10] < xup[10]) && flag);

                                            } while ((j[11] < xup[11]) && flag);

                                        } while ((j[12] < xup[12]) && flag);

                                    } while ((j[13] < xup[13]) && flag);

                                } while ((j[14] < xup[14]) && flag);

                            } while ((j[15] < xup[15]) && flag);

                        } while ((j[16] < xup[16]) && flag);

                    } while ((j[17] < xup[17]) && flag);

                } while ((j[18] < xup[18]) && flag);

            } while ((j[19] < xup[19]) && flag);

            linearList2 = new List<KumiawaseItem>();

            if (kumiawases.Count == 0)
            {
                MessageBox.Show(Constant.ERROR_MESSAGE_NOT_AWASE);
                return;
            }
            for (i = 0; i < kumiawases.Count; i++)
                linearList2.Add(kumiawases.ElementAt(i));

            /*
             * システムのソート
             */
            linearList2.Sort(delegate (KumiawaseItem item1, KumiawaseItem item2) { return item1.spec >= item2.spec ? -1 : 1; });

            for (i = 0; i < kumiawases.Count; i++)
            {
                ListViewItem lvi = new ListViewItem();
                lvi.Text = i.ToString();
                lvi.SubItems.Add(string.Format("グループ{0}", i));
                lvi.SubItems.Add(linearList2[i].spec.ToString());
                listRanking2.Items.Add(lvi);
            }
        }
        
        private void readIndividualHistory(object sender, EventArgs e)
        {
            var t = Directory.GetFiles("Downloads\\");
            var files = Directory.GetFiles("Downloads\\").Where(s => s.StartsWith("Downloads\\zulu_"));
            fileListView.Items.Clear();

            foreach (var file in files)
            {
                ListViewItem lvi = new ListViewItem(Path.GetFileName(file));
                lvi.Checked = true;
                Invoke((MethodInvoker)(() => fileListView.Items.Add(lvi)));

            }
            currentDownload = fileListView.Items.Count;
        }

        /*
         * システムリストを表示
         */
        private void addTradingListView(string trader, string currency, string type, string high, string worst, string count, string profit, string spec, float maxLog, float maxWorst)
        {
            ListViewItem lvi = new ListViewItem(trader);
            lvi.SubItems.Add(currency);
            lvi.SubItems.Add(type);
            lvi.SubItems.Add(high);
            lvi.SubItems.Add(worst);
            lvi.SubItems.Add(count);
            lvi.SubItems.Add(profit);
            lvi.SubItems.Add(spec);
            lvi.SubItems.Add("0.01");
            lvi.SubItems.Add(maxLog.ToString());
            lvi.SubItems.Add(maxWorst.ToString());
            lvi.SubItems.Add("");

            lvi.Checked = true;
            Invoke((MethodInvoker)(() => listTrading.Items.Add(lvi)));
        }

        private void SetPanelEnabledProperty(bool isEnabled)
        {
            // InvokeRequired is used to manage the case the UI is modified
            // from another thread that the UI thread
            if (this.InvokeRequired)
            {
                this.Invoke(new MethodInvoker(() => this.SetPanelEnabledProperty(isEnabled)));
            }
            else
            {
                this.Enabled = isEnabled;
            }
        }

        private void listTrading_MouseUp(object sender, MouseEventArgs e)
        {
            ListViewHitTestInfo i = listTrading.HitTest(e.X, e.Y);
            SelectedLSI = i.SubItem;

            if (SelectedLSI == null)
                return;
            int col = i.Item.SubItems.IndexOf(i.SubItem);
            if (col != 9 && col != 10)
            {
                SelectedLSI = null;
                return;
            }
                

            int border = 0;
            switch (listTrading.BorderStyle)
            {
                case BorderStyle.FixedSingle:
                    border = 1;
                    break;
                case BorderStyle.Fixed3D:
                    border = 2;
                    break;
            }

            int CellWidth = SelectedLSI.Bounds.Width;
            int CellHeight = SelectedLSI.Bounds.Height;
            int CellLeft = border + listTrading.Left + i.SubItem.Bounds.Left;
            int CellTop = listTrading.Top + i.SubItem.Bounds.Top;
            // First Column
            CellWidth = listTrading.Columns[col].Width;

            TxtEdit.Location = new Point(CellLeft, CellTop);
            TxtEdit.Size = new Size(CellWidth, CellHeight);
            TxtEdit.Visible = true;
            TxtEdit.BringToFront();
            TxtEdit.Text = i.SubItem.Text;
            TxtEdit.Select();
            TxtEdit.SelectAll();
        }

        private void HideTextEditor()
        {
            TxtEdit.Visible = false;
            if (SelectedLSI != null)
            {
                SelectedLSI.Text = TxtEdit.Text;
            }
                
            SelectedLSI = null;
            TxtEdit.Text = "";
        }

        private void TxtEdit_Leave(object sender, EventArgs e)
        {
            HideTextEditor();
        }

        private void TxtEdit_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Return)
                HideTextEditor();
        }

        private void listTrading_Scroll(object sender, ScrollEventArgs e)
        {
            HideTextEditor();
        }

        private void fileListView_MouseClick(object sender, MouseEventArgs e)
        {
            
        }

        /*
         * ランキング１のリストのクリーク。
         */
        private void listRanking_MouseDown(object sender, MouseEventArgs e)
        {
            /*
             * システムリストのは初期化
             */
            foreach (ListViewItem lvi in listTrading.Items)
            {
                int index = lvi.Index;
                if (lvi.Checked == false)
                    continue;
                lvi.SubItems[9].Text = "0.01";
                lvi.BackColor = default(Color);
            }

            ListViewHitTestInfo info = listRanking.HitTest(e.X, e.Y);
            if (info == null || info.Item == null)
                return;

            int i = listRanking.Items.IndexOf(info.Item);
            KumiawaseItem item = linearList[i];

            int[] savedX = new int[20];

            /*
             * 選択したのシステムの情報を読み取り
             */
            for (i = 0; i < 20; i++)
            {
                savedX[i] = int.Parse(item.strIndex.Substring(i * 2, 2));
            }

            for (int s = 0; s < 20; s++)
            {
                if (s != 0 && c[s] == 0)
                    continue;

                ListViewItem lvi = listTrading.Items[c[s]];
                lvi.SubItems[9].Text = (0.01 * savedX[s]).ToString();
                lvi.BackColor = Color.CadetBlue;
            }
        }

        /*
         * ランキング２のリストのクリーク。
         */
        private void listRank2_MouseDown(object sender, MouseEventArgs e)
        {
            ListViewHitTestInfo info = listRanking2.HitTest(e.X, e.Y);
            if (info == null || info.Item == null)
                return;

            int i = listRanking2.Items.IndexOf(info.Item);
            KumiawaseItem item = linearList2[i];

            int[] savedX = new int[20];

            for (i = 0; i < 20; i++)
            {
                savedX[i] = int.Parse(item.strIndex.Substring(i * 2, 2));
            }

            for (int s = 0; s < 20; s++)
            {
                if (s != 0 && c2[s] == 0)
                    continue;

                ListViewItem lvi = listTrading.Items[c2[s]];
                lvi.SubItems[9].Text = (0.01 * savedX[s]).ToString();
                lvi.BackColor = Color.LightPink;
            }
        }


        private string cleanTraderName(string name)
        {
            return name.Replace('/', ' ');
        }

        /*
         * システムのグループを取引数・性能に応おうじて整列
         */
        private void resortSystems(int minTradingCount)
        {
            systems.Sort(delegate (TradingSystem system1, TradingSystem system2)
            {
                if (system1.count >= minTradingCount && system2.count >= minTradingCount)
                {
                    if (system1.spec > system2.spec)
                        return -1;
                    else if (system1.spec == system2.spec)
                        return 0;
                    else
                        return 1;
                }
                else if (system1.count >= minTradingCount && system2.count < minTradingCount)
                    return -1;
                else if (system2.count >= minTradingCount && system1.count < minTradingCount)
                    return 1;
                else
                {
                    if (system1.spec > system2.spec)
                        return -1;
                    else if (system1.spec == system2.spec)
                        return 0;
                    else
                        return 1;
                }
            });

            listTrading.Items.Clear();

            foreach (TradingSystem system in systems) {
                addTradingListView(system.trader, system.currency, system.type, 
                    system.high.ToString(), system.worst.ToString(), 
                    system.count.ToString(), system.profit.ToString(), 
                    system.spec.ToString(), system.maxLot, system.maxWorst);
            }
            
        }

        private Color GetColor(string raw)
        {
            using (MD5 md5Hash = MD5.Create())
            {
                byte[] data = md5Hash.ComputeHash(Encoding.UTF8.GetBytes(raw));
                return FromHex(BitConverter.ToString(data).Replace("-", string.Empty).Substring(0, 6));
            }
        }

        /*
         * へクス文字列から色情報を取得
         */

        private Color FromHex(string hex)
        {
            return Color.FromArgb(
                int.Parse(hex.Substring(0, 2), System.Globalization.NumberStyles.HexNumber),
                int.Parse(hex.Substring(2, 2), System.Globalization.NumberStyles.HexNumber),
                int.Parse(hex.Substring(4, 2), System.Globalization.NumberStyles.HexNumber));
        }

        private void MainForm_DragDrop(object sender, DragEventArgs e)
        {
            string[] fileList = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            if (fileList.Count() == 1)
            {
                MessageBox.Show(fileList[0]);
            }
        }


        /*
         * 「結合２」の計算
         */ 
        private void btnFinal2_Click(object sender, EventArgs e)
        {
            if (Convert.ToInt32(comboBox1.SelectedValue) == -1)
            {
                MessageBox.Show(Constant.ERROR_MESSAGE_SELECT_COUNT);
                return;
            }

            Thread thread = new Thread(() => kumiawaseWork2());
            thread.Start();
        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            int currentRowDst = 0;
            int currentColDst = 0;
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkBookDst = null;
            Excel.Worksheet worksheetDst = null;
            xlApp.Visible = false;
            string dstFileName = "Downloads//test3.xlsx";
            dstFileName = System.IO.Path.GetFullPath(dstFileName);

            xlWorkBookDst = xlApp.Workbooks.Open(dstFileName, 0, false, 5, "", "", false,
                Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

            worksheetDst = (Excel.Worksheet)xlWorkBookDst.ActiveSheet;

            currentRowDst = worksheetDst.Cells.Find("*", System.Reflection.Missing.Value,
                           System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                           Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                           false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

            Excel.Range firstEmpty = null;
            int i = 1;
            for (i = 2; i < currentRowDst; i++)
            {
                firstEmpty = worksheetDst.get_Range("B" + i);
                if (string.IsNullOrEmpty(firstEmpty.Value2))
                    break;
            }

            firstEmpty.Select();
            Excel.Range dstRng = worksheetDst.get_Range("A2", "B" + currentRowDst.ToString());
            Excel.Range target = dstRng.SpecialCells(Excel.XlCellType.xlCellTypeBlanks, Type.Missing);
            firstEmpty.Select();
            target.Formula = string.Format("=B{0}", i - 1);

            xlWorkBookDst.Save();
            worksheetDst.Columns.AutoFit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheetDst);
            xlWorkBookDst.Close(true);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkBookDst);
            xlApp.UserControl = true;
            xlApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject((object)xlApp);
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            int currentRowDst = 0;
            int currentColDst = 0;
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkBookDst = null;
            Excel.Worksheet worksheetDst = null;
            xlApp.Visible = false;
            string dstFileName = "Downloads//test.xlsx";
            dstFileName = System.IO.Path.GetFullPath(dstFileName);

            xlWorkBookDst = xlApp.Workbooks.Open(dstFileName, 0, false, 5, "", "", false,
                Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

            worksheetDst = (Excel.Worksheet)xlWorkBookDst.ActiveSheet;

            currentRowDst = worksheetDst.Cells.Find("*", System.Reflection.Missing.Value,
                           System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                           Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                           false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

            Excel.Range firstEmpty = null;
            for (int i = 2; i < currentRowDst; i++)
            {
                firstEmpty = worksheetDst.get_Range("A" + i);
                if (string.IsNullOrEmpty(firstEmpty.Value2))
                    break;
            }

            Excel.Range dstRng = worksheetDst.get_Range("A2", "B" + currentRowDst.ToString());
            Excel.Range target = dstRng.SpecialCells(Excel.XlCellType.xlCellTypeBlanks, Excel.XlSpecialCellsValue.xlTextValues);


            firstEmpty.Select();
            target.Formula = "=A2";

            xlWorkBookDst.Save();
            worksheetDst.Columns.AutoFit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheetDst);
            xlWorkBookDst.Close(true);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkBookDst);
            xlApp.UserControl = true;
            xlApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject((object)xlApp);
        }
    }

    class JWTResponse
    {
        [JsonProperty(PropertyName = "access_token")]
        public string accesToken;
        [JsonProperty(PropertyName = "token_type")]
        string tokenType;
        [JsonProperty(PropertyName = "refresh_token")]
        string RefreshToken;
        [JsonProperty(PropertyName = "scope")]
        string Scope;
        [JsonProperty(PropertyName = "jti")]
        string JTI;
    }
}


