using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ZuluAnalyzer
{
    class TradeHistoryClass
    {
        public class TradeHistoryItem
        {
            [JsonProperty(PropertyName = "id")]
            public static int id {get; set;}
            [JsonProperty(PropertyName = "tradeId")]
            public static int tradeId {get; set;}
            [JsonProperty(PropertyName = "providerTicket")]
            public string providerTicket {get; set;}
            [JsonProperty("brokerTicket")]
            public string brokerTicket {get; set;}
            [JsonProperty("providerId")]
            public int providerId {get; set;}
            [JsonProperty("providerName")]
            public string providerName {get; set;}
            [JsonProperty("lots")]
            public float lots {get; set;}
            [JsonProperty("tradeType")]
            public string tradeType {get; set;}
            [JsonProperty("dateOpen")]
            public long dateOpen {get; set;}
            [JsonProperty("dateClosed")]
            public long dateClosed {get; set;}
            [JsonProperty("priceOpen")]
            public float priceOpen {get; set;}
            [JsonProperty("priceClosed")]
            public float priceClosed {get; set;}
            [JsonProperty("maxProfit")]
            public float maxProfit {get; set;}
            [JsonProperty("maxProfitDate")]
            public long maxProfitDate {get; set;}
            [JsonProperty("worstDrawdown")]
            public float worstDrawdown {get; set;}
            [JsonProperty("maxDrawdownDate")]
            public long maxDrawdownDate {get; set;}
            [JsonProperty("pips")]
            public float pips {get; set;}
            [JsonProperty("totalPips")]
            public float totalPips {get; set;}
            [JsonProperty("grossPnl")]
            public float grossPnl {get; set;}
            [JsonProperty("netPnl")]
            public float netPnl {get; set;}
            [JsonProperty("interest")]
            public float interest {get; set;}
            [JsonProperty("commission")]
            public float commission {get; set;}
            [JsonProperty("existsAsLiveTrade")]
            public bool existsAsLiveTrade {get; set;}
            [JsonProperty("hasEconomicEvent")]
            public bool hasEconomicEvent {get; set;}
            [JsonProperty("statusId")]
            public int statusId {get; set;}
            [JsonProperty("valid")]
            public bool valid {get; set;}
            [JsonProperty("amount")]
            public float amount {get; set;}
            [JsonProperty("currencyId")]
            public int currencyId {get; set;}
            [JsonProperty("transactionCurrency")]
            public string transactionCurrency {get; set;}
            [JsonProperty("currency")]
            public string currency {get; set;}
            [JsonProperty("pipMultiplier")]
            public int pipMultiplier {get; set;}
            [JsonProperty("totalAccumulatedPips")]
            public float totalAccumulatedPips {get; set;}
            [JsonProperty("totalAccumulatedPnl")]
            public float totalAccumulatedPnl {get; set;}
        }

        public class TraderContent
        {
            [JsonProperty("content")]
            public List<TradeHistoryItem> content {get; set;}
            [JsonProperty("last")]
            public bool last {get; set;}
            [JsonProperty("totalPages")]
            public int totalPages {get; set;}
            [JsonProperty("numberOfElements")]
            public int numberOfElements {get; set;}
            [JsonProperty("first")]
            public bool first {get; set;}
        }
    }
}
