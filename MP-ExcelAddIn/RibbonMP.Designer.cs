namespace MP_ExcelAddIn
{
    partial class RibbonMP : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonMP ()
            : base (Globals.Factory.GetRibbonFactory ())
        {
            InitializeComponent ();
        }

        /// <summary> 
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose (bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose ();
            }
            base.Dispose (disposing);
        }

        #region Код, автоматически созданный конструктором компонентов

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent ()
        {
            this.MPtab = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.button1 = this.Factory.CreateRibbonButton();
            this.MPtab.SuspendLayout();
            this.group2.SuspendLayout();
            this.SuspendLayout();
            // 
            // MPtab
            // 
            this.MPtab.Groups.Add(this.group1);
            this.MPtab.Groups.Add(this.group2);
            this.MPtab.Label = "Метран Проект";
            this.MPtab.Name = "MPtab";
            // 
            // group1
            // 
            this.group1.Label = "Общие";
            this.group1.Name = "group1";
            // 
            // group2
            // 
            this.group2.Items.Add(this.button1);
            this.group2.Label = "Трудозатраты";
            this.group2.Name = "group2";
            // 
            // button1
            // 
            this.button1.Label = "Расчет стоимости этапа";
            this.button1.Name = "button1";
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // RibbonMP
            // 
            this.Name = "RibbonMP";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.MPtab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonMP_Load);
            this.MPtab.ResumeLayout(false);
            this.MPtab.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab MPtab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonMP RibbonMP
        {
            get { return this.GetRibbon<RibbonMP> (); }
        }
    }
}
