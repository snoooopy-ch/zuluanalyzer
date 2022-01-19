using System;
using System.Collections.Generic;
using System.Security.Cryptography;
using System.Linq;
using System.Threading;
using System.Text;
using System.IO;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Net.Http;
using System.Threading.Tasks;
using System.Net;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Management;

namespace ZuluAnalyzer
{
    // Global enum and structs

    class Constant
    {
        public static string APP_GUID = "{9F6F0AC4-BAA1-45fd-A64F-20000226AAAA}";

        public static string g_strSiteURL = "https://hotmail.com/trade";
        public static string g_strSiteLoginURL1 = "https://japan.zulutrade.com/login";
        public static string g_strSiteLoginURL2 = "https://japan.zulutrade.com/login";
        public static string g_strSiteTradeURL = "https://japan.zulutrade.com/traders";
        public static string g_strSiteGetToken = "https://japan.zulutrade.com/zulutrade-client/auth/jwt";
        public static string g_strSiteTrader = "https://japan.zulutrade.com/zulutrade-client/v2/api/user/providers/performance/search";
        public static string g_strSiteTraderDetail = "https://japan.zulutrade.com/zulutrade-client/trading/api/providers/{0}/trades/history?timeframe=10000&page=0&size=10&sort=dateOpen,asc";
        public static string g_strSiteExportxls = "https://japan.zulutrade.com/zulutrade-client/trading/api/providers/{0}/trades/history/export?timeframe=10000&cid=&exportType=xlsx";
        public static string g_strExchangeUrl = "https://financialmodelingprep.com/api/v3/forex";

        // Messages
        public static string ERROR_MESSAGE = "エラーメッセージ:";
        public static string ERROR_MESSAGE_RECENT_PERIOD = "最近取引期間のエラー";
        public static string ERROR_MESSAGE_MIN_PERIOD = "裁定取引期間のエラー";
        public static string ERROR_MESSAGE_MIN_FOLLOWER = "最小追従数";
        public static string ERROR_MESSAGE_DOWNLOAD = "タウンロードの数";
        public static string PLEASE_WAIT = "いまとりひきりれきの解釈なかですね。少々お待ちください。";
        public static string ERROR_MESSAGE_SELECT_COUNT = "組み合わせの数を入力してください。";
        public static string ERROR_NO_MATCH = "条件に合う組み合わせが見つかりません。";
        public static string ERROR_MESSAGE_BLOCKED = "トレード{0}がブラックされました。";
        public static string ERROR_MESSAGE_NO_FILE = "計算させた履歴ファイルがありません。";
        public static string ERROR_MESSAGE_INVALID_FILE = "履歴ファイルの読み込みで形式エラーが発生しました。";
        public static string ERROR_MESSAGE_SUCCESS_READ = "履歴ファイルの読み込みが成功しました。";
        public static string ERROR_MESSAGE_FAILED_READ = "履歴ファイルの読み込みで形式エラーが発生しました。";
        public static string ERROR_MESSAGE_NOT_AWASE = "その条件に合る組み合わせがみつかりませんでした。";

        // Login
        public static string SITE_LOGIN_FAILED = "ログインに失敗しました。";
        public static string SITE_LOGIN_STARTED = "ログインが始まりました。";
        public static string SITE_LOGIN_RESTARTED = "***** ログインが再開されました。*****";
        public static string SITE_LOGIN_REQUEST_SUCCESS1 = "ログイン要請に成功しました。(段階-1)";
        public static string SITE_LOGIN_REQUEST_SUCCESS2 = "ログイン要請に成功しました。(段階-2)";
        public static string SITE_LOGIN_REQUEST_SUCCESS3 = "ログイン要請に成功しました。(段階-3)";
        public static string SITE_LOGIN_SUCCESS = "ログインに成功しました。";
        public static string SITE_TRADER_DOWNLOADING = "トレードのダウンロード中です。";

        // Fetch Traders
        public static string SITE_FETCH_TRADER_FAILED = "Traderのリストの呼び出しに失敗しました。";
        public static string SITE_FETCH_TRADER_SUCCESS = "Traderのリストの呼び出しに成功しました。";

        // Fetch TraderHistory
        public static string SITE_FETCH_TRADEHISTORY_FAILED = "「{0}:{1}」の取引履歴の呼び出しに失敗しました。";
        public static string SITE_FATCH_HISTORY_COUNT = "{0}個のファイルが検索されました。";

        public static string ZULUANALYSE_STEP1_SUCCESS = "取引履歴の解釈の１段階の成功";
        public static string ZULUANALYSE_STEP1_FAILED = "取引履歴の解釈の１段階の失敗";

        public static string ZULUANALYSE_STEP2_SUCCESS = "取引履歴の解釈の2段階の成功";
        public static string ZULUANALYSE_STEP2_FAILED = "取引履歴の解釈の2段階の失敗";

        public static string ZULUANALYSE_STEP3_SUCCESS = "取引履歴の解釈の3段階の成功";
        public static string ZULUANALYSE_STEP3_FAILED = "取引履歴の解釈の3段階の失敗";

        public static string ZULUANALYSE_STEP4_SUCCESS = "取引履歴の解釈の4段階の成功";
        public static string ZULUANALYSE_STEP4_FAILED = "取引履歴の解釈の4段階の失敗";

        public static string ZULUANALYSE_STEP5_SUCCESS = "取引履歴の解釈の5段階の成功";
        public static string ZULUANALYSE_STEP5_FAILED = "取引履歴の解釈の5段階の失敗";

        public static string ZULUANALYSE_STEP0_SUCCESS = "相場のダウンロードに成功しました。";
        public static string ZULUANALYSE_STEP0_FAILED = "相場のダウンロードに失敗しました。";
        public static string ZULUANALYSE_EXCHANGE_RATE = "Financial Modeling Prep {0} 相場です。";

        public static string ZULUANALYSE_EXCHAGE_FAILED = "{0}の相場計算が失敗しました。";

        // Export xls
        public static string SITE_EXPORT_XLS_FAILED = "「{0}:{1}」の取引履歴のファイルの呼び出しに失敗しました。";

        // Setting
        public static string PROXY_TEST_FAILED = "プロキシテストに失敗しました。";

        // Open HttpFlag
        public static int OPENED_HTTP = 0;
        public static int FIRST_HTTP = 1;
    }

    public class IniControl
    {
        [DllImport("kernel32", CharSet = CharSet.Unicode)]
        public static extern bool WritePrivateProfileString(string lpAppName, string lpKeyName, string lpString, string lpFileName);

        [DllImport("kernel32", CharSet = CharSet.Unicode)]
        public static extern int GetPrivateProfileInt(string lpAppName, string lpKeyName, int nDefault, string lpFileName);

        [DllImport("kernel32", CharSet = CharSet.Unicode)]
        public static extern int GetPrivateProfileString(string lpAppName, string lpKeyName, string lpDefault, StringBuilder lpReturnedString, int nSize, String lpFileName);

        public static string Base64Encode(string plainText)
        {
            var plainTextBytes = System.Text.Encoding.UTF8.GetBytes(plainText);
            return System.Convert.ToBase64String(plainTextBytes);
        }

        public static string Base64Decode(string cipherText)
        {
            byte[] data = System.Convert.FromBase64String(cipherText);
            return System.Text.UTF8Encoding.UTF8.GetString(data);
        }


        public static void WriteIntValue(string lpAppName, string lpKeyName, int nValue, string lpFileName)
        {
            WritePrivateProfileString(lpAppName, lpKeyName, nValue.ToString(), lpFileName);
        }

        public static void WriteBoolValue(string lpAppName, string lpKeyName, bool bValue, string lpFileName)
        {
            WritePrivateProfileString(lpAppName, lpKeyName, bValue ? "1" : "0", lpFileName);
        }

        public static void WriteStringValue(string lpAppName, string lpKeyName, string lpValue, string lpFileName)
        {
            WritePrivateProfileString(lpAppName, lpKeyName, Base64Encode(lpValue), lpFileName);
        }

        public static string GetStringValue(string lpAppName, string lpKeyName, string lpDefault, int nSize, string lpFileName)
        {
            StringBuilder sbBuffer = new StringBuilder(nSize);

            GetPrivateProfileString(lpAppName, lpKeyName, lpDefault, sbBuffer, nSize, lpFileName);

            return Base64Decode(sbBuffer.ToString());
        }

        public static int GetIntValue(string lpAppName, string lpKeyName, int nDefault, string lpFileName)
        {
            return GetPrivateProfileInt(lpAppName, lpKeyName, nDefault, lpFileName);
        }

        public static bool GetBoolValue(string lpAppName, string lpKeyName, int nDefault, string lpFileName)
        {
            int nValue = GetPrivateProfileInt(lpAppName, lpKeyName, nDefault, lpFileName);

            return nValue == 1 ? true : false;
        }

        public static double GetDoubleValue(string lpAppName, string lpKeyName, double dDefault, string lpFileName)
        {
            double dResult = 0;

            int nSize = 0x10;
            StringBuilder sbBuffer = new StringBuilder(nSize);
            string lpDefault = dDefault.ToString();
            GetPrivateProfileString(lpAppName, lpKeyName, lpDefault, sbBuffer, nSize, lpFileName);
            Double.TryParse(sbBuffer.ToString(), out dResult);
            return dResult;
        }
    }

    class CGlobalVar
    {
        public static string CONFIG_FILE_NAME = "ZuluAnalyzer.ini";
        public static string g_strConfigPath;

        // Download Setting
        public static int g_nRecentTradingPeriod = 0;
        public static int g_nMinimumTradingPeriod = 0;
        public static int g_nMinimumFollowingCount = 0;
        public static int g_nDownloadCount = 0;

        // Login Info
        public static string g_strLoginID = "DM500225312";
        public static string g_strPassword = "POuUgo9U";

        // Setting
        public static bool g_bUseProxy = false;
        public static string g_strProxyIP = "";
        public static int g_nProxyPort = 0;
        public static string g_strProxyID = "";
        public static string g_strProxyPass = "";

        public static void ReadConfig()
        {
            g_strConfigPath = Path.Combine(Application.StartupPath, CONFIG_FILE_NAME);

            g_nRecentTradingPeriod = IniControl.GetIntValue("DownloadSetting", "RecentTradingPeriod", 2, g_strConfigPath);
            g_nMinimumTradingPeriod = IniControl.GetIntValue("DownloadSetting", "MinimumTradingPeriod", 6, g_strConfigPath);
            g_nMinimumFollowingCount = IniControl.GetIntValue("DownloadSetting", "MinimumFollowingCount", 300, g_strConfigPath);
            g_nDownloadCount = IniControl.GetIntValue("DownloadSetting", "DownloadCount", 1, g_strConfigPath);

            g_strLoginID = IniControl.GetStringValue("LoginInfo", "LoginID", "", 256, g_strConfigPath);
            g_strPassword = IniControl.GetStringValue("LoginInfo", "Password", "", 256, g_strConfigPath);

            g_bUseProxy = IniControl.GetBoolValue("Setting", "UseProxy", 0, g_strConfigPath);
            g_strProxyIP = IniControl.GetStringValue("Setting", "ProxyIP", "", 100, g_strConfigPath);
            g_nProxyPort = IniControl.GetIntValue("Setting", "ProxyPort", 0, g_strConfigPath);
            g_strProxyID = IniControl.GetStringValue("Setting", "ProxyID", "", 100, g_strConfigPath);
            g_strProxyPass = IniControl.GetStringValue("Setting", "ProxyPass", "", 100, g_strConfigPath);
        }

        public static void WriteConfig()
        {
            IniControl.WriteIntValue("DownloadSetting", "RecentTradingPeriod", g_nRecentTradingPeriod, g_strConfigPath);
            IniControl.WriteIntValue("DownloadSetting", "MinimumTradingPeriod", g_nMinimumTradingPeriod, g_strConfigPath);
            IniControl.WriteIntValue("DownloadSetting", "MinimumFollowingCount", g_nMinimumFollowingCount, g_strConfigPath);
            IniControl.WriteIntValue("DownloadSetting", "DownloadCount", g_nDownloadCount, g_strConfigPath);

            IniControl.WriteStringValue("LoginInfo", "LoginID", g_strLoginID, g_strConfigPath);
            IniControl.WriteStringValue("LoginInfo", "Password", g_strPassword, g_strConfigPath);

            IniControl.WriteBoolValue("Setting", "UseProxy", g_bUseProxy, g_strConfigPath);
            IniControl.WriteStringValue("Setting", "ProxyIP", g_strProxyIP, g_strConfigPath);
            IniControl.WriteIntValue("Setting", "ProxyPort", g_nProxyPort, g_strConfigPath);
            IniControl.WriteStringValue("Setting", "ProxyID", g_strProxyID, g_strConfigPath);
            IniControl.WriteStringValue("Setting", "ProxyPass", g_strProxyPass, g_strConfigPath);
        }

        public static string EncodeURIComponent(Dictionary<string, object> rgData)
        {
            string _result = String.Join("&", rgData.Select((x) => String.Format("{0}={1}", x.Key, x.Value)));

            _result = System.Net.WebUtility.UrlEncode(_result)
                        .Replace("+", "%20").Replace("%21", "!")
                        .Replace("%27", "'").Replace("%28", "(")
                        .Replace("%29", ")").Replace("%26", "&")
                        .Replace("%3D", "=").Replace("%7E", "~");
            return _result;
        }

        public static bool isSubSequence(string str1,
                  string str2, int m, int n)
        {

            // Base Cases 
            if (m == 0)
                return true;
            if (n == 0)
                return false;

            // If last characters of two strings 
            // are matching 
            if (str1[m - 1] == str2[n - 1])
                return isSubSequence(str1, str2,
                                        m - 1, n - 1);

            // If last characters are not matching 
            return isSubSequence(str1, str2, m, n - 1);
        }
    }
}
