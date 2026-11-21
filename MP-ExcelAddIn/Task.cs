using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Threading.Tasks;

namespace MP_ExcelAddIn
{
    public class Task
    {
        public string TaskName { get; set; }            //Наименование задачи
        public Double TaskCost { get; set; }            //Стоимость
        public string TaskCurrency { get; set; }        //Валюта
        public Double TaskPrice { get; set; }           //Цена
        public string TaskDescription { get; set; }     //Примечание

        public Task() {}
        public Task (string taskName, double taskCost, string taskCurrency, string taskDescription, List<Currency> currencyList, String projectCurrency)
        {
            TaskName = taskName;
            TaskCost = taskCost;
            TaskCurrency = taskCurrency;
            TaskDescription = taskDescription;
            CalculateCurrency (projectCurrency, currencyList);
        }

        internal void CalculateCurrency (string _projectCurrency, List<Currency> _currencyList)
        {
            if (_projectCurrency == "Рубль" || _projectCurrency == "рубль")
            {
                if(TaskCurrency == "руб.")
                {
                    TaskPrice = TaskCost;
                }
                else if (TaskCurrency == "долл.")
                {
                    TaskPrice = TaskCost * _currencyList [_currencyList.FindIndex (x => x.Name == "Доллар США")].Rate;
                }    
                else if (TaskCurrency == "евро")
                {
                    TaskPrice = TaskCost * _currencyList [_currencyList.FindIndex (x => x.Name == "Евро")].Rate;
                }
                else if (TaskCurrency == "юань")
                {
                    TaskPrice = TaskCost * _currencyList [_currencyList.FindIndex (x => x.Name == "Юань")].Rate;
                }
            }
            else if (_projectCurrency == "Доллар США" || _projectCurrency == "доллар США")
            {
                
                if (TaskCurrency == "руб.")
                {
                    TaskPrice = TaskCost/_currencyList [_currencyList.FindIndex (x => x.Name == "Доллар США")].Rate;
                }
                else if (TaskCurrency == "долл.")
                {
                    TaskPrice = TaskCost;
                }
                else if (TaskCurrency == "евро")
                {
                    TaskPrice = TaskCost * (_currencyList [_currencyList.FindIndex (x => x.Name == "Евро")].Rate/
                                            _currencyList [_currencyList.FindIndex (x => x.Name == "Доллар США")].Rate);
                }
                else if (TaskCurrency == "юань")
                {
                    TaskPrice = TaskCost * _currencyList [_currencyList.FindIndex (x => x.Name == "Юань")].Rate/
                                           _currencyList [_currencyList.FindIndex (x => x.Name == "Доллар США")].Rate;
                }
                
            }
            else if (_projectCurrency == "Евро" || _projectCurrency == "евро")
            {
                if (TaskCurrency == "руб.")
                {
                    TaskPrice = TaskCost / _currencyList [_currencyList.FindIndex (x => x.Name == "Евро")].Rate;
                }
                else if (TaskCurrency == "долл.")
                {
                    TaskPrice = TaskCost * _currencyList [_currencyList.FindIndex (x => x.Name == "Доллар США")].Rate/
                                           _currencyList [_currencyList.FindIndex (x => x.Name == "Евро")].Rate;
                }
                else if (TaskCurrency == "евро")
                {
                    TaskPrice = TaskCost;
                }
                else if (TaskCurrency == "юань")
                {
                    TaskPrice = TaskCost * _currencyList [_currencyList.FindIndex (x => x.Name == "Юань")].Rate/
                                           _currencyList [_currencyList.FindIndex (x => x.Name == "Евро")].Rate;
                }
            }
            else if (_projectCurrency == "Юань" || _projectCurrency == "юань")
            {
                if (TaskCurrency == "руб.")
                {
                    TaskPrice = TaskCost / _currencyList [_currencyList.FindIndex (x => x.Name == "Юань")].Rate;
                }
                else if (TaskCurrency == "долл.")
                {
                    TaskPrice = TaskCost * _currencyList [_currencyList.FindIndex (x => x.Name == "Доллар США")].Rate/
                                           _currencyList [_currencyList.FindIndex (x => x.Name == "Юань")].Rate;
                }
                else if (TaskCurrency == "евро")
                {
                    TaskPrice = TaskCost * _currencyList [_currencyList.FindIndex (x => x.Name == "Евро")].Rate/
                                           _currencyList [_currencyList.FindIndex (x => x.Name == "Юань")].Rate;
                }
                else if (TaskCurrency == "юань")
                {
                    TaskPrice = TaskCost;
                }
            }
        }
    }
}
