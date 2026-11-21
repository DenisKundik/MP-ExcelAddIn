using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MP_ExcelAddIn
{
    public class TaskSite : TaskOffice
    {
        public int NumberOfBusinessTrips { get; set; }               //Количество командировок
        public Double TaskBusinessTravelAllowances { get; set; }        //Командировочные
        public Double TaskDailyAllowances { get; set; }                 //Суточные
        public Double TaskTransport { get; set; }                       //Транспортные
        public Double TaskAccommodation { get; set; }                   //Проживание
        public Double TaskAmountDailyAllowance {  get; set; }           //Размер суточных
        public int TaskFirstDayOff { get; set; }                       //Первый день выходной
        public int TaskLastDayOff { get; set; }                        //Последний день выходной
        
        public TaskSite (String rateType,
                            int numberOfPeople,
                            Double duration,
                            String calendarWorking,
                            String dayHour,
                            String firstDay,
                            String typeOfWork,
                            String executor,
                            int numberOfBusinessTrips,
                            String taskDescription,
                            int lengthWorkingDay,
                            String workOnWeekends,
                            Double taskAmountDailyAllowance,
                            List<Rate> rateList,
                            List<Currency> currencyList,
                            String projectCurrency,
                            double accommodation,
                            double travel) :
            base(rateType, numberOfPeople, duration, calendarWorking, dayHour, firstDay, typeOfWork, executor, taskDescription, lengthWorkingDay, workOnWeekends, rateList, currencyList, projectCurrency)
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
            NumberOfBusinessTrips = numberOfBusinessTrips;
            FirstLastDayOff (FirstDay,Duration);
            CalculateHours (DayHour, CalendarWorking, WorkOnWeekends, TaskFirstDayOff, TaskLastDayOff);
            CalculateRate (rateList.Find (r => r.ExecutorName == Executor),RateType);
            rate = rateList.Find (r => r.ExecutorName == Executor);
            FOT = RegularHours * rate.SalesSupportRate + OvertimeHours * rate.SalesSupportRate * 2;
            
            TaskDailyAllowances = CalculateTaskDailyAllowances (NumberOfPeople, Duration, NumberOfBusinessTrips, taskAmountDailyAllowance);
            TaskTransport = CalculateTaskTransport (NumberOfPeople, NumberOfBusinessTrips, travel);
            TaskAccommodation = CalculateTaskAccommodation (NumberOfPeople, Duration, NumberOfBusinessTrips, accommodation);
            TaskBusinessTravelAllowances = TaskDailyAllowances + TaskTransport + TaskAccommodation;
            CalculateCurrency (projectCurrency, currencyList);
        }

        internal void CalculateHours (bool _dayHour, bool _calendarWorking, bool _workOnWeekends, int _taskFirstDayOff, int _taskLastDayOff)
        {
            if (_dayHour == true)//Трудозатраты в днях
            {
                NumberOfDaysOff = Weekends ((int) Duration, FirstDay);
                ManDays = NumberOfPeople * Duration * NumberOfBusinessTrips;
                ManHours = ManDays * LengthWorkingDay;
                if (_calendarWorking)//Календарные дни
                {
                    RegularHours = NumberOfPeople * NumberOfBusinessTrips * (Duration - NumberOfDaysOff) * LengthWorkingDay;
                }
                else//Рабочие дни
                {
                    RegularHours = ManHours;
                }

                if (_workOnWeekends)//Работа в выходные
                {
                    OvertimeHours = NumberOfPeople * NumberOfBusinessTrips * NumberOfDaysOff * LengthWorkingDay;
                }
                else
                {
                    if (Duration > 1)
                    {
                        OvertimeHours = NumberOfPeople * NumberOfBusinessTrips * (_taskFirstDayOff + _taskLastDayOff) * LengthWorkingDay;
                    }
                    else
                    {
                        OvertimeHours = NumberOfPeople * NumberOfBusinessTrips * LengthWorkingDay;
                    }
                }
            }
            
        }

        internal void FirstLastDayOff (String _firstDay, double _duration)
        {
            int x = 0;

            if (IndexWeekDay(_firstDay) == 6 || IndexWeekDay (_firstDay) == 7)
            {
                TaskFirstDayOff = 1;
            }
            else
            {
                TaskFirstDayOff = 0;
            }
            x = (IndexWeekDay (_firstDay) + (int)_duration + 5) % 7;
            if (x == 5 || x == 6)
            {
                TaskLastDayOff = 1;
            }
            else
            {
                TaskLastDayOff = 0;
            }
        }

        internal double CalculateTaskDailyAllowances (int _numberOfPeople, double _duration, double _numberOfBusinessTrips, double _taskAmountDailyAllowance)
        {
            return _numberOfPeople * _duration * _numberOfBusinessTrips * _taskAmountDailyAllowance;
        }

        internal double CalculateTaskTransport(int _numberOfPeople, double _numberOfBusinessTrips, double _travel)
        {
            return _numberOfPeople * _numberOfBusinessTrips * _travel;
        }
        internal double CalculateTaskAccommodation (int _numberOfPeople, double _duration, double _numberOfBusinessTrips, double _accommodation)
        {
            return _numberOfBusinessTrips * (_numberOfPeople * (_accommodation * (_duration - 1)));
        }
        internal new void CalculateCurrency (string _projectCurrency, List<Currency> _currencyList)
        {
            if (_projectCurrency == "Рубль" || _projectCurrency == "рубль")
            {
                TaskPrice = Rate + TaskBusinessTravelAllowances;
            }
            else if (_projectCurrency == "Доллар США" || _projectCurrency == "доллар США")
            {
                TaskPrice = (Rate + TaskBusinessTravelAllowances) * _currencyList [_currencyList.FindIndex (x => x.Name == "Доллар США")].Rate;
            }
            else if (_projectCurrency == "Евро" || _projectCurrency == "евро")
            {
                TaskPrice = (Rate + TaskBusinessTravelAllowances) * _currencyList [_currencyList.FindIndex (x => x.Name == "Евро")].Rate;
            }
            else if (_projectCurrency == "Юань" || _projectCurrency == "юань")
            {
                TaskPrice = (Rate + TaskBusinessTravelAllowances) * _currencyList [_currencyList.FindIndex (x => x.Name == "Юань")].Rate;
            }
        }

    }
}
