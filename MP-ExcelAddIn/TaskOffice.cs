using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Microsoft.Office.Interop.Excel;

using MP_ExcelAddIn;

namespace MP_ExcelAddIn
{

    public class TaskOffice:Task
    {
        public static Rate rate;
        
        /// <summary>
        /// Для вычисления количества выходных
        /// </summary>
        int [,] we = {{0,1,2,3,4,5,6},
                     {0,0,0,0,0,0,1},
                     {0,0,0,0,0,1,2},
                     {0,0,0,0,1,2,2},
                     {0,0,0,1,2,2,2},
                     {0,0,1,2,2,2,2},
                     {0,1,2,2,2,2,2},
                     {0,1,1,1,1,1,1}};

        
        public String RateType { get; set; }            //Тип ставки
        public int NumberOfPeople { get; set; }         //Количество человек
        public Double Duration { get; set; }            //Продолжительность
        public bool CalendarWorking { get; set; }       //Если true - календарные дни, false - рабочие дни
        public bool DayHour { get; set; }               //Если true - дней, false - часов
        public String FirstDay { get; set; }            //День недели начала периода
        public String TypeOfWork { get; set; }          //Вид работ
        public String Executor { get; set; }            //Исполнитель
        public Double ManDays { get; set; }             //Чел./дней
        public Double ManHours { get; set; }            //Чел./часов
        public Double RegularHours { get; set; }        //Нормальных часов
        public Double OvertimeHours { get; set; }       //Сверхурочных часов
        public Double Rate {  get; set; }               //Стоимость по ставкам
        public Double FOT {  get; set; }                //ФОТ
        public Double Overhead { get; set; }            //Накладные расходы
        public int LengthWorkingDay { get; set; }       //Продолжительность рабочего дня
        public bool WorkOnWeekends {  get; set; }       //Работа в выходыне
        public int NumberOfDaysOff { get; set; }        //Количество выходных дней в периоде

        
        public TaskOffice (String rateType,
                          int numberOfPeople,
                          Double duration,
                          String calendarWorking,
                          String dayHour,
                          String firstDay,
                          String typeOfWork,
                          String executor,
                          String taskDescription,
                          int lengthWorkingDay,
                          String workOnWeekends,
                          List<Rate> rateList,
                          List<Currency> currencyList,
                          String projectCurrency)
        {
            RateType = rateType;
            NumberOfPeople = numberOfPeople;
            Duration = duration;
            CalendarWorking = SetCalendarWorking (calendarWorking);
            DayHour = SetDayHour (dayHour);
            FirstDay = firstDay;
            TypeOfWork = typeOfWork;
            Executor = executor;
            TaskDescription = taskDescription;
            LengthWorkingDay = lengthWorkingDay;
            WorkOnWeekends = SetWorkOnWeekends (workOnWeekends);
            CalculateHours (DayHour, CalendarWorking, WorkOnWeekends);
            CalculateRate (rateList.Find (r => r.ExecutorName == Executor), RateType);
            rate = rateList.Find (r => r.ExecutorName == Executor);
            FOT = RegularHours * rate.SalesSupportRate + OvertimeHours * rate.SalesSupportRate * 2;
            CalculateCurrency (projectCurrency, currencyList);
        }

        /// <summary>
        /// Подсчет количества выходных в заданном периоде, начинающимся с определенного дня недели
        /// </summary>
        /// <param name="day">Количество дней</param>
        /// <param name="weekday">День недели</param>
        /// <returns></returns>
        internal int Weekends (int day, String weekday)
        {
            int weekends=0;         //Выходных в указанном диапазоне
            int w1 = 0;             //Выходных в полных неделях
            int w2 = 0;             //Выходных в оставщихся днях
            int weeknumber = 0;     //Число полных недель в периоде
            int restdays = 0;       //остатко дней
            int iwd = 0;            //Индекс дня недели

            iwd = IndexWeekDay (weekday);
            weeknumber = day / 7;
            restdays = day - weeknumber * 7;
            w1 = weeknumber * 2;
            w2 = we [iwd, restdays];
            weekends = w1 + w2;
            return weekends;
        }

        /// <summary>
        /// Определение индекса дня недели
        /// </summary>
        /// <param name="s">День недели (сокращение)</param>
        /// <returns></returns>
        internal int IndexWeekDay (String s)
        {
            int i = 0;

            switch (s)
            {
                case "Пн":
                    i = 1;
                    break;
                case "Вт":
                    i = 2;
                    break;
                case "Ср":
                    i = 3;
                    break;
                case "Чт":
                    i = 4;
                    break;
                case "Пт":
                    i = 5;
                    break;
                case "Сб":
                    i = 6;
                    break;
                case "Вс":
                    i = 7;
                    break;
            }
            return i;
        }

        /// <summary>
        /// true - календарные дни
        /// </summary>
        /// <param name="_calendarWorking"></param>
        /// <returns></returns>
        internal bool SetCalendarWorking(string _calendarWorking)
        {
            if (_calendarWorking == "К" || _calendarWorking == "к")
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        internal bool SetDayHour(string _dayHour)
        {
            if (_dayHour.Contains ("Час"))
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        internal bool SetWorkOnWeekends(string _workOnWeekends)
        {
            if (_workOnWeekends == "Да")
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        internal void CalculateHours(bool _dayHour, bool _calendarWorking, bool _workOnWeekends)
        {
            if (_dayHour == true)//Трудозатраты в днях
            {
                NumberOfDaysOff = Weekends ((int) Duration, FirstDay);
                ManDays = NumberOfPeople * Duration;
                ManHours = ManDays * LengthWorkingDay;
                if (_calendarWorking)//Календарные дни
                {
                    RegularHours = NumberOfPeople * (Duration - NumberOfDaysOff) * LengthWorkingDay;
                }
                else//Рабочие дни
                {
                    RegularHours = ManHours;
                }

                if (_workOnWeekends)//Работа в выходные
                {
                    OvertimeHours = NumberOfPeople * NumberOfDaysOff * LengthWorkingDay;
                }
                else
                {
                    OvertimeHours = 0;
                }
            }
            else//Трудозатраты в часах
            {
                NumberOfDaysOff = 0;
                ManDays = NumberOfPeople * Duration / LengthWorkingDay;
                ManHours = NumberOfPeople * Duration;
                CalendarWorking = false;
            }
        }

        internal void CalculateRate(Rate _rate, string _rateType)
        {
            if (_rateType == "П")//Выбор типа ставок
            {
                Rate = RegularHours * _rate.ProjectRate + OvertimeHours * _rate.ProjectOvertimeRate;
                Overhead = (RegularHours * _rate.ProjectRate + OvertimeHours * _rate.ProjectOvertimeRate) -
                    (RegularHours * _rate.SalesSupportRate + OvertimeHours * _rate.SalesSupportRate * 2);
            }
            else if (_rateType == "З")
            {
                Rate = RegularHours * _rate.CustomerRate + OvertimeHours * _rate.CustomerOvertimeRate;
                Overhead = (RegularHours * _rate.CustomerRate + OvertimeHours * _rate.CustomerOvertimeRate) -
                    (RegularHours * _rate.SalesSupportRate + OvertimeHours * _rate.SalesSupportRate * 2);
            }
            else
            {
                Rate = RegularHours * _rate.SalesSupportRate + OvertimeHours * _rate.SalesSupportRate * 2;
                Overhead = 0;
            }
        }

        internal new void CalculateCurrency(string _projectCurrency, List<Currency> _currencyList)
        {
            if (_projectCurrency == "Рубль" || _projectCurrency == "рубль")
            {
                TaskPrice = Rate;
            }
            else if (_projectCurrency == "Доллар США" || _projectCurrency == "доллар США")
            {
                TaskPrice = Rate * _currencyList [_currencyList.FindIndex (x => x.Name == "Доллар США")].Rate;
            }
            else if (_projectCurrency == "Евро" || _projectCurrency == "евро")
            {
                TaskPrice = Rate * _currencyList [_currencyList.FindIndex (x => x.Name == "Евро")].Rate;
            }
            else if (_projectCurrency == "Юань" || _projectCurrency == "юань")
            {
                TaskPrice = Rate * _currencyList [_currencyList.FindIndex (x => x.Name == "Юань")].Rate;
            }
        }
    }
}
