using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;

namespace MP_ExcelAddIn
{
    public partial class RibbonMP
    {
        List<TaskOffice> taskOfficeList = new List<TaskOffice>();
        List<TaskSite> taskSiteList = new List<TaskSite>();
        List<Task> taskSubcontractor = new List<Task>();
        List<Task> taskOther = new List<Task>();
        public String ProjectCurrency;       //Валюта проекта
        public Double Accommodation;         //Проживание
        public Double Travel;                //Проезд
        
        public List<Rate> RateList = new List<Rate>();
        public List<Currency> CurrencyList = new List<Currency> ();

        private void RibbonMP_Load (object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click (object sender, RibbonControlEventArgs e)
        {
            CalculationOfLaborCosts ();
        }

        /// <summary>
        /// Расчет стоимости трудозатрат
        /// </summary>
        public void CalculationOfLaborCosts ()
        {
            #region init variable

            int OfficeRowStart = 0;
            int OfficeRowEnd = 0;
            int SiteRowStart = 0;
            int SiteRowEnd = 0;
            int SubcontractorRowStart = 0;
            int SubcontractorRowEnd = 0;
            int OtherCostRowStart = 0;
            int OtherCostRowEnd = 0;

            Range range = Globals.ThisAddIn.Application.ActiveCell.Worksheet.UsedRange;                         //используемая область на листе

            int usedRowsNum = Globals.ThisAddIn.Application.ActiveCell.Worksheet.UsedRange.Rows.Count;          //количество строк в используемой области на листе
            int usedColumnsNum = Globals.ThisAddIn.Application.ActiveCell.Worksheet.UsedRange.Columns.Count;    //количество стобцов в используемой области на листе

            
            #endregion

            if (range.Count != 0)
            {
                CurrencyList.Clear ();
                RateList.Clear ();
                taskOfficeList.Clear ();
                taskSiteList.Clear ();
                taskSubcontractor.Clear ();
                taskOther.Clear ();
                GetSettings ();
                GetRate ();

                for (int r = 1; r <= usedRowsNum; r++)//Поиск номеров строк
                {
                    try
                    {
                        if (range.Cells [r, 2].Value2 == "Работа в офисе")
                            OfficeRowStart = r + 2;
                        if (range.Cells [r, 2].Value2 == "Работа в офисе, всего:")
                            OfficeRowEnd = r - 1;
                        if (range.Cells [r, 2].Value2 == "Командировки на площадку")
                            SiteRowStart = r + 2;
                        if (range.Cells [r, 2].Value2 == "Работа в командировке, всего:")
                            SiteRowEnd = r - 1;
                        if (range.Cells [r, 2].Value2 == "Субподрядчик")
                            SubcontractorRowStart = r + 2;
                        if (range.Cells [r, 2].Value2 == "Субподрядчик, всего:")
                            SubcontractorRowEnd = r - 1;
                        if (range.Cells [r, 2].Value2 == "Прочие затраты")
                            OtherCostRowStart = r + 1;
                        if (range.Cells [r, 2].Value2 == "Прочие, всего:")
                            OtherCostRowEnd = r - 1;
                    }
                    catch (Exception ex) { }
                }
                
                for (int r = OfficeRowStart; r <= OfficeRowEnd; r++)
                {
                    try
                    {
                        if (range.Cells [r, 2].Value2 != null && range.Cells [r, 5].Value2 != null && range.Cells [r, 10].Value2 != null)
                        {
                            taskOfficeList.Add (new TaskOffice ((String) range.Cells [r, 1].Value2,
                            (int) range.Cells [r, 2].Value2,
                            (Double) range.Cells [r, 5].Value2,
                            (String) range.Cells [r, 6].Value2,
                            (String) range.Cells [r, 7].Value2,
                            (String) range.Cells [r, 8].Value2,
                            (String) range.Cells [r, 9].Value2,
                            (String) range.Cells [r, 10].Value2,
                            (String) range.Cells [r, 19].Value2,
                            (int) range.Cells [r, 22].Value2,
                            (String) range.Cells [r, 23].Value2,
                            RateList, CurrencyList, ProjectCurrency));

                        }

                    }
                    catch (Exception e) { }
                }
                for (int r = 0; r<= taskOfficeList.Count; r++)
                {
                    try
                    {
                        range.Cells [OfficeRowStart + r, 12].Value2 = taskOfficeList[r].ManDays;
                        range.Cells [OfficeRowStart + r, 13].Value2 = taskOfficeList [r].ManHours;
                        range.Cells [OfficeRowStart + r, 14].Value2 = taskOfficeList [r].RegularHours;
                        range.Cells [OfficeRowStart + r, 15].Value2 = taskOfficeList [r].OvertimeHours;
                        range.Cells [OfficeRowStart + r, 16].Value2 = taskOfficeList [r].Rate;
                        range.Cells [OfficeRowStart + r, 18].Value2 = taskOfficeList [r].TaskPrice;
                        range.Cells [OfficeRowStart + r, 20].Value2 = taskOfficeList [r].FOT;
                        range.Cells [OfficeRowStart + r, 21].Value2 = taskOfficeList [r].Overhead;
                        range.Cells [OfficeRowStart + r, 24].Value2 = taskOfficeList [r].NumberOfDaysOff;
                    }
                    catch(Exception e) { }
                }
                for (int r = taskOfficeList.Count; r <= OfficeRowEnd - OfficeRowStart; r++)
                {
                    try
                    {
                        range.Cells [OfficeRowStart + r, 12].Value2 = 0;
                        range.Cells [OfficeRowStart + r, 13].Value2 = 0;
                        range.Cells [OfficeRowStart + r, 14].Value2 = 0;
                        range.Cells [OfficeRowStart + r, 15].Value2 = 0;
                        range.Cells [OfficeRowStart + r, 16].Value2 = 0;
                        range.Cells [OfficeRowStart + r, 18].Value2 = 0;
                        range.Cells [OfficeRowStart + r, 20].Value2 = 0;
                        range.Cells [OfficeRowStart + r, 21].Value2 = 0;
                        range.Cells [OfficeRowStart + r, 24].Value2 = 0;
                    }
                    catch (Exception e) { }
                }

                for (int r = SiteRowStart; r <= SiteRowEnd; r++)
                {
                    try
                    {
                        if (range.Cells [r, 2].Value2 != null && range.Cells [r, 5].Value2 != null && range.Cells [r, 10].Value2 != null)
                        {
                            taskSiteList.Add (new TaskSite ((String) range.Cells [r, 1].Value2,
                            (int) range.Cells [r, 2].Value2,
                            (Double) range.Cells [r, 5].Value2,
                            (String) range.Cells [r, 6].Value2,
                            (String) range.Cells [r, 7].Value2,
                            (String) range.Cells [r, 8].Value2,
                            (String) range.Cells [r, 9].Value2,
                            (String) range.Cells [r, 10].Value2,
                            (int) range.Cells [r, 11].Value2,
                            (String) range.Cells [r, 19].Value2,
                            (int) range.Cells [r, 22].Value2,
                            (String) range.Cells [r, 23].Value2,
                            (Double) range.Cells [r,27].Value2,
                            RateList, CurrencyList, ProjectCurrency,Accommodation, Travel));

                        }

                    }
                    catch (Exception e) { }
                }
                for (int r = 0; r <= taskSiteList.Count; r++)
                {
                    try
                    {
                        range.Cells [SiteRowStart + r, 12].Value2 = taskSiteList [r].ManDays;
                        range.Cells [SiteRowStart + r, 13].Value2 = taskSiteList [r].ManHours;
                        range.Cells [SiteRowStart + r, 14].Value2 = taskSiteList [r].RegularHours;
                        range.Cells [SiteRowStart + r, 15].Value2 = taskSiteList [r].OvertimeHours;
                        range.Cells [SiteRowStart + r, 16].Value2 = taskSiteList [r].Rate;
                        range.Cells [SiteRowStart + r, 17].Value2 = taskSiteList [r].TaskBusinessTravelAllowances;
                        range.Cells [SiteRowStart + r, 18].Value2 = taskSiteList [r].TaskPrice;
                        range.Cells [SiteRowStart + r, 20].Value2 = taskSiteList [r].FOT;
                        range.Cells [SiteRowStart + r, 21].Value2 = taskSiteList [r].Overhead;
                        range.Cells [SiteRowStart + r, 24].Value2 = taskSiteList [r].NumberOfDaysOff;
                        range.Cells [SiteRowStart + r, 25].Value2 = taskSiteList [r].TaskFirstDayOff;
                        range.Cells [SiteRowStart + r, 26].Value2 = taskSiteList [r].TaskLastDayOff;
                        range.Cells [SiteRowStart + r, 28].Value2 = taskSiteList [r].TaskDailyAllowances;
                        range.Cells [SiteRowStart + r, 29].Value2 = taskSiteList [r].TaskTransport;
                        range.Cells [SiteRowStart + r, 30].Value2 = taskSiteList [r].TaskAccommodation;
                    }
                    catch (Exception e) { }
                }
                for (int r = taskSiteList.Count; r <= SiteRowEnd - SiteRowStart; r++)
                {
                    try
                    {
                        range.Cells [SiteRowStart + r, 12].Value2 = 0;
                        range.Cells [SiteRowStart + r, 13].Value2 = 0;
                        range.Cells [SiteRowStart + r, 14].Value2 = 0;
                        range.Cells [SiteRowStart + r, 15].Value2 = 0;
                        range.Cells [SiteRowStart + r, 16].Value2 = 0;
                        range.Cells [SiteRowStart + r, 17].Value2 = 0;
                        range.Cells [SiteRowStart + r, 18].Value2 = 0;
                        range.Cells [SiteRowStart + r, 20].Value2 = 0;
                        range.Cells [SiteRowStart + r, 21].Value2 = 0;
                        range.Cells [SiteRowStart + r, 24].Value2 = 0;
                        range.Cells [SiteRowStart + r, 25].Value2 = 0;
                        range.Cells [SiteRowStart + r, 26].Value2 = 0;
                        range.Cells [SiteRowStart + r, 28].Value2 = 0;
                        range.Cells [SiteRowStart + r, 29].Value2 = 0;
                        range.Cells [SiteRowStart + r, 30].Value2 = 0;
                    }
                    catch (Exception e) { }
                }

                for (int r = SubcontractorRowStart; r <= SubcontractorRowEnd; r++)
                {
                    try
                    {
                        if (range.Cells [r, 2].Value2 != null && range.Cells [r, 11].Value2 != null)
                        {
                            taskSubcontractor.Add (new Task ((String) range.Cells [r, 2].Value2,
                            (Double) range.Cells [r, 11].Value2,
                            (String) range.Cells [r, 12].Value2,
                            (String) range.Cells [r, 19].Value2,
                            CurrencyList, ProjectCurrency));

                        }

                    }
                    catch (Exception e) { }
                }
                for (int r = 0; r <= taskSubcontractor.Count; ++r)
                {
                    try
                    {
                        range.Cells [SubcontractorRowStart + r, 18].Value2 = taskSubcontractor[r].TaskPrice;
                    }
                    catch (Exception e) { }
                }
                for (int r = taskSubcontractor.Count; r <= SubcontractorRowEnd - SubcontractorRowStart; r++)
                {
                    range.Cells [SubcontractorRowStart + r, 2].Value2 = null;
                    range.Cells [SubcontractorRowStart + r, 11].Value2 = null;
                    range.Cells [SubcontractorRowStart + r, 12].Value2 = null;
                    range.Cells [SubcontractorRowStart + r, 18].Value2 = null;
                    range.Cells [SubcontractorRowStart + r, 19].Value2 = null;
                }

                for (int r = OtherCostRowStart; r <= OtherCostRowEnd; r++)
                {
                    try
                    {
                        if (range.Cells [r, 2].Value2 != null && range.Cells [r, 11].Value2 != null)
                        {
                            taskOther.Add (new Task ((String) range.Cells [r, 2].Value2,
                            (Double) range.Cells [r, 11].Value2,
                            (String) range.Cells [r, 12].Value2,
                            (String) range.Cells [r, 19].Value2,
                            CurrencyList, ProjectCurrency));

                        }

                    }
                    catch (Exception e) { }
                }
                for (int r = 0; r <= taskOther.Count; ++r)
                {
                    try
                    {
                        range.Cells [OtherCostRowStart + r, 18].Value2 = taskOther [r].TaskPrice;
                    }
                    catch (Exception e) { }
                }
                for (int r = taskOther.Count; r <= OtherCostRowEnd - OtherCostRowStart; r++)
                {
                    range.Cells [OtherCostRowStart + r, 2].Value2 = null;
                    range.Cells [OtherCostRowStart + r, 11].Value2 = null;
                    range.Cells [OtherCostRowStart + r, 12].Value2 = null;
                    range.Cells [OtherCostRowStart + r, 18].Value2 = null;
                    range.Cells [OtherCostRowStart + r, 19].Value2 = null;
                }
            }
        }

        /// <summary>
        /// Получение настроек с листа "Общие"
        /// </summary>
        internal void GetSettings ()
        {
            Worksheet settingsSheet = Globals.ThisAddIn.Application.Sheets["Общие"];//ссылка на лист настроек
            Range rangeSettingsSheet = settingsSheet.UsedRange;//ссылка на диапазон используемых ячеек на листе настроек
            if (rangeSettingsSheet.Cells [1, 6].Value2 != null)
            {
                ProjectCurrency = (string)rangeSettingsSheet.Cells [1, 6].Value2;
            }
            if (rangeSettingsSheet.Cells [7, 6].Value2 != null)
            {
                Accommodation = (double)rangeSettingsSheet.Cells [7, 6].Value2;
            }
            if (rangeSettingsSheet.Cells [8, 6].Value2 != null)
            {
                Travel = (double)rangeSettingsSheet.Cells [8, 6].Value2;
            }
            

            CurrencyList.Add (new Currency ((string) rangeSettingsSheet.Cells [3, 1].Value2,
                                            (string) rangeSettingsSheet.Cells [3, 2].Value2,
                                            (string) rangeSettingsSheet.Cells [3, 3].Value2,
                                            (double) rangeSettingsSheet.Cells [3, 4].Value2,
                                            (string) rangeSettingsSheet.Cells [3, 5].Value2,
                                            (double) rangeSettingsSheet.Cells [3, 6].Value2,
                                            (string) rangeSettingsSheet.Cells [3, 7].Value2));
            CurrencyList.Add (new Currency ((string) rangeSettingsSheet.Cells [4, 1].Value2,
                                            (string) rangeSettingsSheet.Cells [4, 2].Value2,
                                            (string) rangeSettingsSheet.Cells [4, 3].Value2,
                                            (double) rangeSettingsSheet.Cells [4, 4].Value2,
                                            (string) rangeSettingsSheet.Cells [4, 5].Value2,
                                            (double) rangeSettingsSheet.Cells [4, 6].Value2,
                                            (string) rangeSettingsSheet.Cells [4, 7].Value2));
            CurrencyList.Add (new Currency ((string) rangeSettingsSheet.Cells [5, 1].Value2,
                                            (string) rangeSettingsSheet.Cells [5, 2].Value2,
                                            (string) rangeSettingsSheet.Cells [5, 3].Value2,
                                            (double) rangeSettingsSheet.Cells [5, 4].Value2,
                                            (string) rangeSettingsSheet.Cells [5, 5].Value2,
                                            (double) rangeSettingsSheet.Cells [5, 6].Value2,
                                            (string) rangeSettingsSheet.Cells [5, 7].Value2));
            CurrencyList.Add (new Currency ((string) rangeSettingsSheet.Cells [6, 1].Value2,
                                            (string) rangeSettingsSheet.Cells [6, 2].Value2,
                                            (string) rangeSettingsSheet.Cells [6, 3].Value2,
                                            (double) rangeSettingsSheet.Cells [6, 4].Value2,
                                            (string) rangeSettingsSheet.Cells [6, 5].Value2,
                                            (double) rangeSettingsSheet.Cells [6, 6].Value2,
                                            (string) rangeSettingsSheet.Cells [6, 7].Value2));

        }

        /// <summary>
        /// Получение ставок с листа "Ставки"
        /// </summary>
        internal void GetRate ()
        {
            Worksheet rateSheet = Globals.ThisAddIn.Application.Sheets["Ставки"];//ссылка на лист ставок
            Range rangeRateSheet = rateSheet.UsedRange;//ссылка на диапазон используемых ячеек на листе настроек
            int usedRowsNum = rangeRateSheet.Rows.Count;          //количество строк в используемой области на листе
            int usedColumnsNum = rangeRateSheet.Columns.Count;    //количество стобцов в используемой области на листе
            if (rangeRateSheet.Count != 0)
            {
                for (int r = 3; r <= usedRowsNum; r++)
                {
                    try
                    {
                        if (rangeRateSheet.Cells [r, 2] != null)
                        {
                            RateList.Add (new Rate ((String) rangeRateSheet.Cells [r, 2].Value2,
                                                      (Double) rangeRateSheet.Cells [r, 3].Value2,
                                                      (Double) rangeRateSheet.Cells [r, 4].Value2,
                                                      (Double) rangeRateSheet.Cells [r, 5].Value2,
                                                      (Double) rangeRateSheet.Cells [r, 6].Value2,
                                                      (Double) rangeRateSheet.Cells [r, 7].Value2));
                        }
                    }
                    catch (Exception e)
                    {

                    }
                }
            }

        }


    }
}
