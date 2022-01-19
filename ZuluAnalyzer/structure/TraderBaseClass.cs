using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ZuluAnalyzer
{
    class TraderBaseClass
    {
        public class TraderProfile {
            [JsonProperty(PropertyName = "zuluAccountId")]
            public int zuluAccountId {get; set;}
            [JsonProperty(PropertyName = "brokerAccountId")]
            public int brokerAccountId {get; set;}
            [JsonProperty(PropertyName = "name")]
            public string name {get; set;}
            [JsonProperty(PropertyName = "strategyDesc")]
            public string strategyDesc {get; set;}
            [JsonProperty(PropertyName = "strategyDescApproved")]
            public bool strategyDescApproved {get; set;}
            [JsonProperty(PropertyName = "pageVisitis")]
            public int pageVisitis {get; set;}
            [JsonProperty(PropertyName = "publicTradeHistory")]
            public bool publicTradeHistory {get; set;}
            [JsonProperty(PropertyName = "photoApproved")]
            public bool photoApproved {get; set;}
            [JsonProperty(PropertyName = "active")]
            public bool active {get; set;}
            [JsonProperty(PropertyName = "brokerId")]
            public int brokerId {get; set;}
            [JsonProperty(PropertyName = "brokerName")]
            public string brokerName {get; set;}
            [JsonProperty(PropertyName = "feedGroupId")]
            public int feedGroupId {get; set;}
            [JsonProperty(PropertyName = "valid")]
            public bool valid {get; set;}
            [JsonProperty(PropertyName = "loggedInRecently")]
            public bool loggedInRecently {get; set;}
            [JsonProperty(PropertyName = "baseCurrencyName")]
            public string baseCurrencyName {get; set;}
            [JsonProperty(PropertyName = "baseCurrencySymbol")]
            public string baseCurrencySymbol {get; set;}
            [JsonProperty(PropertyName = "demo")]
            public bool demo {get; set;}
            [JsonProperty(PropertyName = "id")]
            public int id {get; set;}
            [JsonProperty(PropertyName = "traderPublic")]
            public bool traderPublic {get; set;}
        }

        public class traderOverallStats
        {
            [JsonProperty(PropertyName = "amountFollowing")]
            public float amountFollowing {get; set;}
            [JsonProperty(PropertyName = "amountFollowingNew")]
            public float amountFollowingNew {get; set;}
            [JsonProperty(PropertyName = "averageSlippage")]
            public float averageSlippage {get; set;}
            [JsonProperty(PropertyName = "avgDrawdown")]
            public float avgDrawdown {get; set;}
            [JsonProperty(PropertyName = "avgWeeklyBestTrade")]
            public float avgWeeklyBestTrade {get; set;}
            [JsonProperty(PropertyName = "avgWeeklyWorstTrade")]
            public float avgWeeklyWorstTrade {get; set;}
            [JsonProperty(PropertyName = "bestTrade")]
            public float bestTrade {get; set;}
            [JsonProperty(PropertyName = "followers")]
            public int followers {get; set;}
            [JsonProperty(PropertyName = "liveFollowers")]
            public float liveFollowers {get; set;}
            [JsonProperty(PropertyName = "worstTrade")]
            public float worstTrade {get; set;}
            [JsonProperty(PropertyName = "liveFollowersPnl")]
            public float liveFollowersPnl {get; set;}
            [JsonProperty(PropertyName = "totalFollowerProfit")]
            public float totalFollowerProfit {get; set;}
            [JsonProperty(PropertyName = "providerCurrencies")]
            public string providerCurrencies {get; set;}
            [JsonProperty(PropertyName = "weeks")]
            public int weeks {get; set;}
            [JsonProperty(PropertyName = "roiAnnualized")]
            public float roiAnnualized {get; set;}
            [JsonProperty(PropertyName = "roiProfit")]
            public float roiProfit {get; set;}
            [JsonProperty(PropertyName = "correlationPercent")]
            public float correlationPercent {get; set;}
            [JsonProperty(PropertyName = "hedgingPercent")]
            public float hedgingPercent {get; set;}
            [JsonProperty(PropertyName = "lastOpenTradeDate")]
            public long lastOpenTradeDate {get; set;}
            public long firstOpenTradeDate { get; set; }
            [JsonProperty(PropertyName = "lastUpdatedDate")]
            public long lastUpdatedDate {get; set;}
            [JsonProperty(PropertyName = "zuluRank")]
            public int zuluRank {get; set;}
            [JsonProperty(PropertyName = "overallDrawDownPercent")]
            public float overallDrawDownPercent {get; set;}
            [JsonProperty(PropertyName = "initialInvestmentTierForActualRoiInMoney")]
            public float initialInvestmentTierForActualRoiInMoney {get; set;}
            [JsonProperty(PropertyName = "rorBasedRoi")]
            public double rorBasedRoi {get; set;}
            [JsonProperty(PropertyName = "initialBalance")]
            public float initialBalance {get; set;}
        }

        public class TraderRate
        {
            [JsonProperty(PropertyName = "overall")]
            public float overall {get; set;}
            [JsonProperty(PropertyName = "overallPercent")]
            public float overallPercent {get; set;}
            [JsonProperty(PropertyName = "count")]
            public int count {get; set;}
        }

        public class CurrencyStatsItem
        {
            [JsonProperty(PropertyName = "currencyId")]
            public int currencyId {get; set;}
            [JsonProperty(PropertyName = "currencyName")]
            public string currencyName {get; set;}
            [JsonProperty(PropertyName = "pips")]
            public float pops {get; set;}
            [JsonProperty(PropertyName = "totalCurrencyCount")]
            public int totalCurrencyCount {get; set;}
            [JsonProperty(PropertyName = "currencyWinCount")]
            public int currencyWinCount {get; set;}
            [JsonProperty(PropertyName = "currencyWinPercent")]
            public float currencyWinPercent {get; set;}
            [JsonProperty(PropertyName = "totalCurrencyBuyCount")]
            public int totalCurrencyBuyCount {get; set;}
            [JsonProperty(PropertyName = "totalCurrencySellCount")]
            public int totalCurrencySellCount {get; set;}
        }

        public class TimeFrameStatsItem
        {
            [JsonProperty(PropertyName = "timeFrameId")]
            public int timeFrameId {get; set;}
            [JsonProperty(PropertyName = "profit")]
            public float profit {get; set;}
            [JsonProperty(PropertyName = "totalProfit")]
            public float totalProfit {get; set;}
            [JsonProperty(PropertyName = "totalProfitMoney")]
            public float totalProfitMoney {get; set;}
            [JsonProperty(PropertyName = "trades")]
            public int trades {get; set;}
            [JsonProperty(PropertyName = "maxOpenTrades")]
            public int maxOpenTrades {get; set;}
            [JsonProperty(PropertyName = "avgPipsPerTrade")]
            public float avgPipsPerTrade {get; set;}
            [JsonProperty(PropertyName = "avgPnlPerTrade")]
            public float avgPnlPerTrade {get; set;}
            [JsonProperty(PropertyName = "avgWorstTrade")]
            public float avgWorstTrade {get; set;}
            [JsonProperty(PropertyName = "winTrades")]
            public float winTrades {get; set;}
            [JsonProperty(PropertyName = "winTradesCount")]
            public int winTradesCount {get; set;}
            [JsonProperty(PropertyName = "winTradesInMoney")]
            public float winTradesInMoney {get; set;}
            [JsonProperty(PropertyName = "winTradesCountInMoney")]
            public int winTradesCountInMoney {get; set;}
            [JsonProperty(PropertyName = "nme")]
            public float nme {get; set;}
            [JsonProperty(PropertyName = "maxDrawDown")]
            public float maxDrawDown {get; set;}
            [JsonProperty(PropertyName = "maxDrawDownPercent")]
            public float maxDrawDownPercent {get; set;}
            [JsonProperty(PropertyName = "worstDrawDownPercent")]
            public float worstDrawDownPercent {get; set;}
            [JsonProperty(PropertyName = "overallDrawDown")]
            public float overallDrawDown {get; set;}
            [JsonProperty(PropertyName = "overallDrawDownMoney")]
            public float overallDrawDownMoney {get; set;}
            [JsonProperty(PropertyName = "overallDrawDownPercent")]
            public float overallDrawDownPercent {get; set;}
            [JsonProperty(PropertyName = "totalClosedPipsAtMaxDrawdownTime")]
            public float totalClosedPipsAtMaxDrawdownTime {get; set;}
            [JsonProperty(PropertyName = "avgTradeSeconds")]
            public float avgTradeSeconds {get; set;}
            [JsonProperty(PropertyName = "liveFollowersPnl")]
            public float liveFollowersPnl {get; set;}
            [JsonProperty(PropertyName = "totalFollowerProfit")]
            public float totalFollowerProfit {get; set;}
            [JsonProperty(PropertyName = "minLotSize")]
            public string minLotSize {get; set;}
            [JsonProperty(PropertyName = "statsInMoneyValidStartDate")]
            public long statsInMoneyValidStartDate {get; set;}
            [JsonProperty(PropertyName = "rorBasedRoi")]
            public double rorBasedRoi {get; set;}
            [JsonProperty(PropertyName = "timeFrameInitialEquity")]
            public float timeFrameInitialEquity {get; set;}
            [JsonProperty(PropertyName = "currencyStats")]
            public List<CurrencyStatsItem> currencyStats {get; set;}
        }

        public class TimeFrameStats
        {
            [JsonProperty(PropertyName = "30")]
            public TimeFrameStatsItem i30 {get; set;}
        }

        public class ResultItem
        {
            [JsonProperty(PropertyName = "trader")]
            public Trader trader {get; set;}
        }

        public class TraderItem
        {
            [JsonProperty(PropertyName = "providerId")]
            public int providerId {get; set;}
            [JsonProperty(PropertyName = "profile")]
            public TraderProfile profile {get; set;}
            [JsonProperty(PropertyName = "overallStats")]
            public traderOverallStats overallStats {get; set;}
            [JsonProperty(PropertyName = "rate")]
            public TraderRate rate {get; set;}
            [JsonProperty(PropertyName = "currencyStats")]
            public List<CurrencyStatsItem> currencyStats {get; set;}
            [JsonProperty(PropertyName = "openCurrencyStats")]
            public List<CurrencyStatsItem> openCurrencyStats {get; set;}
            [JsonProperty(PropertyName = "timeframeStats")]
            public TimeFrameStats timeframeStats {get; set;}
        }

        public class Trader
        {
            [JsonProperty(PropertyName = "trader")]
            public TraderItem trader {get; set;}
        }

        public class TraderResponse
        {
            [JsonProperty(PropertyName = "result")]
            public List<Trader> result {get; set;}
        }
    }
}
