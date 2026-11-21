using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MP_ExcelAddIn
{
    public class Currency
    {
        public string Name { get; set; }            //Наименование валюты
        public string ShortName { get; set; }       //Сокращение
        public string Code { get; set; }            // ISO-код валюты (RUB, USD, EUR, CNY)
        public double DCode { get; set; }           //Цифровой код валютя
        public string Symbol { get; set; }          //Символ валюты
        public double Rate { get; set; }            // Курс
        public string CurrencyRate { get; set; }    //Валюта курса

        public Currency (string name, string shortName, string code, double dcode, string symbol, double rate, string currencyRate)
        {
            this.Name = name;
            this.ShortName = shortName;
            this.Code = code;
            this.DCode = dcode;
            this.Symbol = symbol;
            this.Rate = rate;
            this.CurrencyRate = currencyRate;
        }
    }
}
