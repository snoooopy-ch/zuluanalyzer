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
    public partial class Form1 : Form
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
        List<string> currentList = null;

        ListViewItem.ListViewSubItem SelectedLSI;

        int[] c = { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };

        private string[] map = new string[]
        {
            "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"
        };

        public string getExcelColumnLetter(int number)
        {
            return map[number];
        }
        public Form1()
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
            //pp.Add(new ComboItem() { id = 15, name = "15" });
            //pp.Add(new ComboItem() { id = 20, name = "20" });

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
            //else if (type == LogType.Combine)
            //    AddCombineLog(log + "\r\n");

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
                //else if (type == LogType.Combine)
                //    AddCombineLog(strErrMsg);
            }
        }

        private void btnStartDownload_Click(object sender, EventArgs e)
        {
            btnStartDownload.Enabled = false;
            clearHistoryFiles();

            Thread thread = new Thread(() => DoLogin(Constant.OPENED_HTTP));
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

        private Result DoLogin(int flag)
        {
            currentDownload = 0;
            if (!checkInputValidate())
                return Result.RETRY;

            if (!downloadExchangeRate())
                return Result.FAILURE;

            int nStartPos = 0, nEndPos = 0;
            try
            {
                if (flag == 0)
                    AddLog(Constant.SITE_LOGIN_STARTED);

                // Login - Step 1
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
                do
                {
                    url = Constant.g_strSiteTrader;
                    http_request.setURL(url);
                    http_request.setSendMode(HTTP_SEND_MODE.HTTP_POST);
                    http_request.setAutoRedirect(true);
                    http_request.appendCustomHeader("authorization: Bearer " + strToken);

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
                foreach (var trader in allTraders.result)
                {
                    if (trader.trader.profile.name.Equals("GoldKeyFX"))
                    {
                        MessageBox.Show("1");
                    }

                    Invoke((MethodInvoker)(() => txtDLFinishedCount.Text = string.Format("{0}人中{1}人検査、成功：{2}", allTraders.result.Count, checkCount, currentDownload)));

                    if (checkCount == 390)
                    {
                        int j = 0;
                        j += 1;
                    }
                    checkCount++;
                    //AddLog(checkCount.ToString());


                    url = string.Format(Constant.g_strSiteTraderDetail, trader.trader.providerId);
                    http_request.setURL(url);
                    http_request.setSendMode(HTTP_SEND_MODE.HTTP_GET);
                    http_request.setAutoRedirect(true);

                    if (checkCount == 300)
                        Thread.Sleep(2000);
                    if (!http_request.sendRequest(false))
                    {
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
            if (currentList.Contains(trader.trader.profile.baseCurrencyName))
            {
                Color color = GetColor(trader.trader.profile.baseCurrencyName);
                lvi.SubItems[5].ForeColor = color;
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

            if (!(blackers.Count() == 1 && blackers[0].Equals("")) && blackers.Any(trader.trader.profile.name.Contains))
            {
                AddLog(string.Format(Constant.ERROR_MESSAGE_BLOCKED, trader.trader.profile.name));
                return false;
            }
                

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
            try
            {
                foreach (ListViewItem lvi in fileListView.Items)
                {
                    if (!lvi.Checked)
                        continue;

                    string fileName = "Downloads//" + lvi.Text;
                    string validFileName = Path.GetFileNameWithoutExtension(fileName);
                    string[] segs = validFileName.Split('_');
                    string traderName = "", traderId, baseCurrencyNameSymbol = "";

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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);

                    Invoke((MethodInvoker)(() => progressBar1.Value = progressBar1.Value + 1));
                }
            }
            catch (Exception e)
            {
                return false;
            }
            return true;
        }

        private bool mergeAllHistory()
        {
            bool retval = true;

            try
            {
                int currentRowDst = 0;
                int currentColDst = 0;
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkBookDst = null;
                Excel.Worksheet worksheetDst = null;
                Excel.Workbook xlWorkBookSrc = null;
                Excel.Worksheet worksheetSrc = null;
                xlApp.Visible = false;

                bool historyFileIsExist = false;
                foreach (ListViewItem lvi in fileListView.Items)
                {
                    if (!lvi.Checked)
                        continue;

                    if (!historyFileIsExist)
                    {
                        string srcFileName = "Downloads//" + lvi.Text;
                        string dstFileName = "Downloads//All_履歴.xlsx";
                        System.IO.File.Copy(srcFileName, dstFileName, true);
                        historyFileIsExist = true;

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
                        string srcFileName = "Downloads//" + lvi.Text;
                        srcFileName = System.IO.Path.GetFullPath(srcFileName);
                        xlWorkBookSrc = xlApp.Workbooks.Open(srcFileName, 0, false, 5, "", "", false,
                            Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                        worksheetSrc = (Excel.Worksheet)xlWorkBookSrc.ActiveSheet;

                        int limit = 1000000;
                        int totalRowsSrc = worksheetSrc.Cells.Find("*", System.Reflection.Missing.Value,
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
                        else
                        {
                            break;
                        }
                        xlWorkBookDst.Save();

                        worksheetDst.Name = "All_履歴";

                        System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheetSrc);
                        xlWorkBookSrc.Close(true);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkBookSrc);
                    }
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheetDst);
                xlWorkBookDst.Close(true);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkBookDst);
                xlApp.UserControl = true;
                xlApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);

                progressBar1.Value += 1;
            }
            catch (Exception e)
            {
                AddLog(e.Message);
                retval = false;
            }
            finally
            {
                
            }
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
                worksheetDst.get_Range("O1", "O" + currentRowDst.ToString()).PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                worksheetDst.get_Range("O1", "O" + currentRowDst.ToString()).NumberFormat = "¥#;[Red]-¥#";
                range1.Clear();
                worksheetDst.Calculate();

                System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheetDst);
                xlWorkBookDst.Close(true);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkBookDst);
                xlApp.UserControl = true;
                xlApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);

                progressBar1.Value += 1;
            }
            catch (Exception e)
            {
                AddLog(e.Message);
                retval = false;
            }
            finally
            {

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
                r1.Formula = "=IF(N2=0,0,ABS((O2)/N2))/F2*0.01";

                worksheetDst.Cells[1, 18].value = "計算行②Highest Profit(¥)";
                Excel.Range r2 = worksheetDst.get_Range("R2", "R" + currentRowSrc.ToString());
                r2.Formula = "=K2*Q2/F2*0.01";
                r2.Cells.NumberFormat = "#.00;[Red]-¥#.00";

                worksheetDst.Cells[1, 19].value = "計算行  ③Worst Drawdown(¥)";
                Excel.Range r3 = worksheetDst.get_Range("S2", "S" + currentRowSrc.ToString());
                r3.Formula = "=L2*Q2/F2*0.01";
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
                worksheetDst.get_Range("G2", "G" + currentRowSrc.ToString()).PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                rg = worksheetDst.get_Range("Z2", "Z" + currentRowSrc.ToString());
                rg.Copy();
                worksheetDst.get_Range("H2", "H" + currentRowSrc.ToString()).PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

                r6.Clear();
                r7.Clear();

                //Calculate the formula for whole worksheet
                System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheetSrc);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheetDst);
                xlWorkBookDst.Close(true);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkBookDst);
                xlApp.UserControl = true;
                xlApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);

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

                string str = getExcelColumnLetter(currentColSrc - 1);
                Excel.Range range = worksheetSrc.get_Range("A1", str + currentRowSrc.ToString());
                range.Copy(worksheetDst.get_Range("A1", Missing.Value));

                //var stopwatch = new Stopwatch();
                //stopwatch.Start();

                xlApp.DisplayAlerts = false;

                // 18 --> 11
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

                //stopwatch.Stop();
                //long elapsed_time = stopwatch.ElapsedMilliseconds;
                
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);

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


                Excel.Range dataRange = worksheetSrc.get_Range("A1", "R" + currentRowSrc);
                Excel.PivotCache cache = xlWorkBookDst.PivotCaches().Add(Excel.XlPivotTableSourceType.xlDatabase, dataRange);
                Excel.PivotTables tables = worksheetDst.PivotTables();
                Excel.PivotTable pt = tables.Add(cache, worksheetDst.Range["A1"], "Pivot", defaultArg, defaultArg);

                pt.Format(Excel.XlPivotFormatType.xlPTNone);
                pt.ShowValuesRow = false;

                // Row Fields
                Excel.PivotField fld = ((Excel.PivotField)pt.PivotFields("System ID"));
                fld.Orientation = Excel.XlPivotFieldOrientation.xlRowField;
                //                  fld.LayoutForm = Excel.XlLayoutFormType.xlOutline;
                fld.LayoutForm = Excel.XlLayoutFormType.xlTabular;
                fld.LayoutCompactRow = true;
                fld.LayoutSubtotalLocation = Excel.XlSubtototalLocationType.xlAtTop;
                fld.set_Subtotals(1, false);
                fld = ((Excel.PivotField)pt.PivotFields("Currency"));
                fld.Orientation = Excel.XlPivotFieldOrientation.xlRowField;
                //                fld.LayoutForm = Excel.XlLayoutFormType.xlOutline;
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);

                AddLog(Constant.ZULUANALYSE_STEP3_SUCCESS);
            }
            catch (Exception e)
            {
                AddLog(Constant.ZULUANALYSE_STEP3_FAILED);
                //xlApp.Quit();
                AddLog(e.Message);
                retVal = false;
            }

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

                // Create a pivot chart using a pivot table as its source
                Excel.PivotTable pivotTable = (Excel.PivotTable)worksheetSrc.PivotTables("Pivot");
                Excel.Shape chartShape = worksheetDst.Shapes.AddChart();
                chartShape.Chart.SetSourceData(pivotTable.TableRange1, Type.Missing);
                xlWorkBookDst.ShowPivotChartActiveFields = true;
                chartShape.Chart.ChartType = Excel.XlChartType.xlColumnStacked;

                // Calculate the formula for whole worksheet
                worksheetDst.Calculate();
                xlWorkBookDst.Save();

                // Fill the blank cell and set filter.
                for (int i = 1; i <= 2; i ++)
                {
                    for (int j = 2; j <= currentRowSrc; j ++)
                    {
                        string strChk = worksheetDst.Cells[j, i].Value2;
                        if (string.IsNullOrEmpty(strChk))
                        {
                            worksheetDst.Cells[j, i].Value2 = worksheetDst.Cells[j - 1, i].Value2;
                        }
                    }
                }
                xlWorkBookDst.Save();

                // Sort by profit as Desc
                dynamic allDataRange = worksheetDst.get_Range("A2", "H" + currentRowSrc.ToString());
                allDataRange.Sort(allDataRange.Columns[8], Excel.XlSortOrder.xlDescending);
                xlWorkBookDst.Save();

                // Set Filters
                allDataRange = worksheetDst.UsedRange;
                allDataRange.AutoFilter(1, "<>", Excel.XlAutoFilterOperator.xlFilterValues, Type.Missing, true);
                Excel.Range filteredRange = allDataRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible, Excel.XlSpecialCellsValue.xlTextValues);
                xlWorkBookDst.Save();

                // Update the listTrading
                for (int i = 2; i <= currentRowSrc; i++)
                {
                    string trader = worksheetDst.Cells[i, 1].Value2.ToString();
                    string currency = worksheetDst.Cells[i, 2].Value2.ToString();
                    string type = worksheetDst.Cells[i, 3].Value2.ToString();
                    string high = worksheetDst.Cells[i, 4].Value2.ToString();
                    string worst = worksheetDst.Cells[i, 5].Value2.ToString();
                    string count = worksheetDst.Cells[i, 6].Value2.ToString();
                    string profit = worksheetDst.Cells[i, 7].Value2.ToString();
                    string spec = worksheetDst.Cells[i, 8].Value2.ToString();

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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheetSrc);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheetDst);
                xlWorkBookDst.Close(true);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkBookDst);
                xlApp.UserControl = true;
                xlApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);

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

            }
            return retval;
        }

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

            return ret;
        }

        private void btnCombine_Click(object sender, EventArgs e)
        {
            
        }

        private void threadwork()
        {
            Invoke((MethodInvoker)(() => processCombineAction()));
        }

        private void processCombineAction()
        {
            Invoke((MethodInvoker)(() => progressBar1.Value = 0));
            if (!contentShiftRightExcel())
                return;

            Invoke((MethodInvoker)(() => progressBar1.Value = currentDownload));
            if (!mergeAllHistory())
                return;

            if (!deleteColumn())
                return;

            Invoke((MethodInvoker)(() => progressBar1.Value = currentDownload + 1));
            if (!zuluAnalyzerStep1())
                return;

            Invoke((MethodInvoker)(() => progressBar1.Value = currentDownload + 2));
            if (!zuluAnalyzerStep2())
                return;

            Invoke((MethodInvoker)(() => progressBar1.Value = currentDownload + 3));
            if (!zuluAnalyzerStep3())
                return;

            Invoke((MethodInvoker)(() => progressBar1.Value = currentDownload + 4));
            if (!zuluAnalyzerStep4())
                return;

            Invoke((MethodInvoker)(() => progressBar1.Value = currentDownload + 5));
            string dstFileName = "Downloads//All_履歴.xlsx";

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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);

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

            Thread thread = new Thread(new ThreadStart(threadwork));
            thread.Start();

            btnCombine.Enabled = false;
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

            
            //await kumiawaseWork();
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

            foreach (ListViewItem lvi in listTrading.Items)
            {
                int value = int.Parse(lvi.SubItems[5].Text);
                if (minTradingCount > value)
                    lvi.Checked = false;
                lvi.BackColor = default(Color);
            }
            
            float[] a = { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
            float[] b = { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
            int[] xup = { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
            Array.Clear(c, 0, 20);
            
            int i = 0;

            foreach (ListViewItem lvi in listTrading.Items)
            {
                int index = lvi.Index;
                if (lvi.Checked == false)
                    continue;

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
                if (i == systemCount)
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
                    width = 25;
                    break;
                case 7://10,000,000
                    width = 10;
                    break;
                case 9://10,077,696
                    width = 6;
                    break;
                case 10://9,765,625
                    width = 5;
                    break;

            }
            int xupstart = xupdasi - width;

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

            j[19] = xupstart;
            do
            {
                j[19]++;
                j[18] = xupstart;
                do
                {
                    j[18]++;
                    j[17] = xupstart;
                    do
                    {
                        j[17]++;
                        j[16] = xupstart;
                        do
                        {
                            j[16]++;
                            j[15] = xupstart;
                            do
                            {
                                j[15]++;
                                j[14] = xupstart;
                                do
                                {
                                    j[14]++;
                                    j[13] = xupstart;
                                    do
                                    {
                                        j[13]++;
                                        j[12] = xupstart;
                                        do
                                        {
                                            j[12]++;
                                            j[11] = xupstart;
                                            do
                                            {
                                                j[11]++;
                                                j[10] = xupstart;
                                                do
                                                {
                                                    j[10]++;
                                                    j[9] = xupstart;
                                                    do
                                                    {
                                                        j[9]++;
                                                        j[8] = xupstart;
                                                        do
                                                        {
                                                            j[8]++;
                                                            j[7] = xupstart;
                                                            do
                                                            {
                                                                j[7]++;
                                                                j[6] = xupstart;
                                                                do
                                                                {
                                                                    j[6]++;
                                                                    j[5] = xupstart;
                                                                    do
                                                                    {
                                                                        j[5]++;
                                                                        j[4] = xupstart;
                                                                        do
                                                                        {
                                                                            j[4]++;
                                                                            j[3] = xupstart;
                                                                            do
                                                                            {
                                                                                j[3]++;
                                                                                j[2] = xupstart;
                                                                                do
                                                                                {
                                                                                    j[2]++;
                                                                                    j[1] = xupstart;
                                                                                    do
                                                                                    {
                                                                                        j[1]++;
                                                                                        j[0] = xupstart;
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
                                                                                                        for (int t = 0; t < 20; t++)
                                                                                                            sp += (systems[t].spec * j[t]);
                                                                                                        item = new KumiawaseItem();
                                                                                                        item.spec = sp;
                                                                                                        item.strIndex = string.Format("{0:00}{1:00}{2:00}{3:00}{4:00}{5:00}{6:00}{7:00}{9:00}{9:00}",
                                                                                                            j[0], j[1], j[2], j[3], j[4], j[5], j[6], j[7], j[8], j[9]);

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
            for (i = 0; i < kumiawases.Count; i++)
                linearList.Add(kumiawases.ElementAt(i));
            

            linearList.Sort(delegate (KumiawaseItem item1, KumiawaseItem item2) { return item1.spec >= item2.spec ? -1 : 1; });

            for (i = 0; i < kumiawases.Count; i ++)
            {
                ListViewItem lvi = new ListViewItem();
                lvi.Text = i.ToString();
                lvi.SubItems.Add(string.Format("グループ{0}", i));
                lvi.SubItems.Add(linearList[i].spec.ToString());
                listRanking.Items.Add(lvi);
            }
        }

        private void button1_Click(object sender, EventArgs e)
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

        private void button2_Click_1(object sender, EventArgs e)
        {
            var t = Directory.GetFiles("Downloads\\");
            var files = Directory.GetFiles("Downloads\\").Where(s => s.StartsWith("Downloads\\zulu_"));

            List<string> currents = new List<string>();
            
            foreach (var file in files)
            {
                string validFileName = Path.GetFileNameWithoutExtension(file);
                string[] segs = validFileName.Split('_');
                string traderName = "", traderId, baseCurrencyNameSymbol;

                if (segs.Length >= 4)
                {
                    int k;
                    for (k = 1; k < segs.Length - 3; k++)
                        traderName += (segs[k] + "_");
                    traderName += segs[k];
                    traderId = segs[k + 1];
                    baseCurrencyNameSymbol = segs[k + 2];

                    currents.Add(baseCurrencyNameSymbol);
                }
            }

            List<string> final = currents.Distinct().ToList();

            foreach(var item in final)
            {
                
            }
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

        private void listTrading_MouseDown(object sender, MouseEventArgs e)
        {
            HideTextEditor();
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

        private void listRanking_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        private void listRanking_MouseClick(object sender, MouseEventArgs e)
        {
            /*
            foreach (ListViewItem lvi in listTrading.Items)
            {
                int index = lvi.Index;
                if (lvi.Checked == false)
                    continue;
                lvi.SubItems[11].Text = "0.01";
                lvi.BackColor = default(Color);
            }


            int i = listRanking.FocusedItem.Index;
            KumiawaseItem item = linearList[i];

            int[] savedX = new int[20];

            for (i = 0; i < 10; i ++)
            {
                //savedX[i] = int.Parse(item.strIndex.Split('-')[i]);
                savedX[i] = int.Parse(item.strIndex.Substring(i * 2, 2));
            }

            for (int s = 0; s < 20; s++)
            {
                if (s != 0 && c[s] == 0)
                    continue;

                ListViewItem lvi = listTrading.Items[c[s]];
                lvi.SubItems[11].Text = (0.01 * savedX[s]).ToString();
                lvi.BackColor = Color.CadetBlue;
            }
            */
        }

        private void fileListView_MouseClick(object sender, MouseEventArgs e)
        {
            
        }

        private void listRanking_MouseDown(object sender, MouseEventArgs e)
        {
            foreach (ListViewItem lvi in listTrading.Items)
            {
                int index = lvi.Index;
                if (lvi.Checked == false)
                    continue;
                lvi.SubItems[11].Text = "0.01";
                lvi.BackColor = default(Color);
            }

            ListViewHitTestInfo info = listRanking.HitTest(e.X, e.Y);
            if (info == null || info.Item == null)
                return;

            int i = listRanking.Items.IndexOf(info.Item);
            //int i = listRanking.FocusedItem.Index;
            KumiawaseItem item = linearList[i];

            int[] savedX = new int[20];

            for (i = 0; i < 10; i++)
            {
                //savedX[i] = int.Parse(item.strIndex.Split('-')[i]);
                savedX[i] = int.Parse(item.strIndex.Substring(i * 2, 2));
            }

            for (int s = 0; s < 20; s++)
            {
                if (s != 0 && c[s] == 0)
                    continue;

                ListViewItem lvi = listTrading.Items[c[s]];
                lvi.SubItems[11].Text = (0.01 * savedX[s]).ToString();
                lvi.BackColor = Color.CadetBlue;
            }
        }

        private string cleanTraderName(string name)
        {
            return name.Replace('/', ' ');
        }

        private void resortSystems(int minTradingCount)
        {
            systems.Sort(delegate (TradingSystem system1, TradingSystem system2)
            {
                if (system1.count >= minTradingCount && system2.count >= minTradingCount)
                    return system1.spec >= system2.spec ? -1 : 1;
                else if (system1.count >= minTradingCount && system2.count < minTradingCount)
                    return -1;
                else if (system2.count >= minTradingCount && system1.count < minTradingCount)
                    return 1;
                else
                    return system1.spec >= system2.spec ? -1 : 1;
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

        private Color FromHex(string hex)
        {
            return Color.FromArgb(
                int.Parse(hex.Substring(0, 2), System.Globalization.NumberStyles.HexNumber),
                int.Parse(hex.Substring(2, 2), System.Globalization.NumberStyles.HexNumber),
                int.Parse(hex.Substring(4, 2), System.Globalization.NumberStyles.HexNumber));
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
