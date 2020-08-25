namespace BOOTH
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            Microsoft.Office.Tools.Ribbon.RibbonSplitButton ImportSplitButton;
            Microsoft.Office.Tools.Ribbon.RibbonSplitButton ProcessFolderButton;
            Microsoft.Office.Tools.Ribbon.RibbonSplitButton TimersButton;
            Microsoft.Office.Tools.Ribbon.RibbonButton ImportPollPadButton;
            this.ImportDS200Button = this.Factory.CreateRibbonButton();
            this.ImportDICEButton = this.Factory.CreateRibbonButton();
            this.ImportVSAPBMDButton = this.Factory.CreateRibbonButton();
            this.ImportDICXButton = this.Factory.CreateRibbonButton();
            this.ProcessDS200Button = this.Factory.CreateRibbonButton();
            this.ProcessVSAPBMDButton = this.Factory.CreateRibbonButton();
            this.ProcessDICEButton = this.Factory.CreateRibbonButton();
            this.ProcessDICXButton = this.Factory.CreateRibbonButton();
            this.CheckinTimerButton = this.Factory.CreateRibbonButton();
            this.CheckinArrivalTimerButton = this.Factory.CreateRibbonButton();
            this.VotingBoothTimerButton = this.Factory.CreateRibbonButton();
            this.BMDTimerButton = this.Factory.CreateRibbonButton();
            this.BallotScanningTimerButton = this.Factory.CreateRibbonButton();
            this.ThroughputArrivalTimerButton = this.Factory.CreateRibbonButton();
            this.BoothTab = this.Factory.CreateRibbonTab();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.ProcessSingleButton = this.Factory.CreateRibbonButton();
            this.ProcessAllButton = this.Factory.CreateRibbonButton();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.group5 = this.Factory.CreateRibbonGroup();
            this.CreateSumStatsButton = this.Factory.CreateRibbonButton();
            this.group6 = this.Factory.CreateRibbonGroup();
            ImportSplitButton = this.Factory.CreateRibbonSplitButton();
            ProcessFolderButton = this.Factory.CreateRibbonSplitButton();
            TimersButton = this.Factory.CreateRibbonSplitButton();
            ImportPollPadButton = this.Factory.CreateRibbonButton();
            this.BoothTab.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.group4.SuspendLayout();
            this.group5.SuspendLayout();
            this.group6.SuspendLayout();
            this.SuspendLayout();
            // 
            // ImportSplitButton
            // 
            ImportSplitButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            ImportSplitButton.Items.Add(this.ImportDS200Button);
            ImportSplitButton.Items.Add(this.ImportDICEButton);
            ImportSplitButton.Items.Add(this.ImportVSAPBMDButton);
            ImportSplitButton.Items.Add(this.ImportDICXButton);
            ImportSplitButton.Label = "Import BMD Log File(s)";
            ImportSplitButton.Name = "ImportSplitButton";
            ImportSplitButton.OfficeImageId = "GetExternalDataFromText";
            // 
            // ImportDS200Button
            // 
            this.ImportDS200Button.Label = "Import DS200 File(s)";
            this.ImportDS200Button.Name = "ImportDS200Button";
            this.ImportDS200Button.ShowImage = true;
            this.ImportDS200Button.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ImportButton_Click);
            // 
            // ImportDICEButton
            // 
            this.ImportDICEButton.Label = "Import Dominion ICE File(s)";
            this.ImportDICEButton.Name = "ImportDICEButton";
            this.ImportDICEButton.ShowImage = true;
            this.ImportDICEButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ImportButton_Click);
            // 
            // ImportVSAPBMDButton
            // 
            this.ImportVSAPBMDButton.Label = "Import VSAP BMD File(s)";
            this.ImportVSAPBMDButton.Name = "ImportVSAPBMDButton";
            this.ImportVSAPBMDButton.ShowImage = true;
            this.ImportVSAPBMDButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ImportButton_Click);
            // 
            // ImportDICXButton
            // 
            this.ImportDICXButton.Label = "Import Dominion ICX File(s)";
            this.ImportDICXButton.Name = "ImportDICXButton";
            this.ImportDICXButton.ShowImage = true;
            this.ImportDICXButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ImportButton_Click);
            // 
            // ProcessFolderButton
            // 
            ProcessFolderButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            ProcessFolderButton.Items.Add(this.ProcessDS200Button);
            ProcessFolderButton.Items.Add(this.ProcessVSAPBMDButton);
            ProcessFolderButton.Items.Add(this.ProcessDICEButton);
            ProcessFolderButton.Items.Add(this.ProcessDICXButton);
            ProcessFolderButton.Label = "Process BMD Log Folder";
            ProcessFolderButton.Name = "ProcessFolderButton";
            ProcessFolderButton.OfficeImageId = "LoadFromQuery";
            ProcessFolderButton.SuperTip = "Open a folder and process all BMD log files in it into a single output file.";
            // 
            // ProcessDS200Button
            // 
            this.ProcessDS200Button.Label = "Process DS200 Folder";
            this.ProcessDS200Button.Name = "ProcessDS200Button";
            this.ProcessDS200Button.ShowImage = true;
            this.ProcessDS200Button.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ProcessFolderMenuButton_Click);
            // 
            // ProcessVSAPBMDButton
            // 
            this.ProcessVSAPBMDButton.Label = "Process VSAP BMD Folder";
            this.ProcessVSAPBMDButton.Name = "ProcessVSAPBMDButton";
            this.ProcessVSAPBMDButton.ShowImage = true;
            this.ProcessVSAPBMDButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ProcessFolderMenuButton_Click);
            // 
            // ProcessDICEButton
            // 
            this.ProcessDICEButton.Label = "Process Dominion IC E Folder";
            this.ProcessDICEButton.Name = "ProcessDICEButton";
            this.ProcessDICEButton.ShowImage = true;
            this.ProcessDICEButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ProcessFolderMenuButton_Click);
            // 
            // ProcessDICXButton
            // 
            this.ProcessDICXButton.Label = "Process Dominion IC X Folder";
            this.ProcessDICXButton.Name = "ProcessDICXButton";
            this.ProcessDICXButton.ShowImage = true;
            this.ProcessDICXButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ProcessFolderMenuButton_Click);
            // 
            // TimersButton
            // 
            TimersButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            TimersButton.Items.Add(this.CheckinTimerButton);
            TimersButton.Items.Add(this.CheckinArrivalTimerButton);
            TimersButton.Items.Add(this.VotingBoothTimerButton);
            TimersButton.Items.Add(this.BMDTimerButton);
            TimersButton.Items.Add(this.BallotScanningTimerButton);
            TimersButton.Items.Add(this.ThroughputArrivalTimerButton);
            TimersButton.Label = "Timers";
            TimersButton.Name = "TimersButton";
            TimersButton.OfficeImageId = "StartTimer";
            // 
            // CheckinTimerButton
            // 
            this.CheckinTimerButton.Label = "Check-In Timer";
            this.CheckinTimerButton.Name = "CheckinTimerButton";
            this.CheckinTimerButton.ShowImage = true;
            // 
            // CheckinArrivalTimerButton
            // 
            this.CheckinArrivalTimerButton.Label = "Check-In Arrival Timer";
            this.CheckinArrivalTimerButton.Name = "CheckinArrivalTimerButton";
            this.CheckinArrivalTimerButton.ShowImage = true;
            // 
            // VotingBoothTimerButton
            // 
            this.VotingBoothTimerButton.Label = "Voting Booth Timer";
            this.VotingBoothTimerButton.Name = "VotingBoothTimerButton";
            this.VotingBoothTimerButton.ShowImage = true;
            // 
            // BMDTimerButton
            // 
            this.BMDTimerButton.Label = "BMD Timer";
            this.BMDTimerButton.Name = "BMDTimerButton";
            this.BMDTimerButton.ShowImage = true;
            // 
            // BallotScanningTimerButton
            // 
            this.BallotScanningTimerButton.Label = "Ballot Scanning Timer";
            this.BallotScanningTimerButton.Name = "BallotScanningTimerButton";
            this.BallotScanningTimerButton.ShowImage = true;
            // 
            // ThroughputArrivalTimerButton
            // 
            this.ThroughputArrivalTimerButton.Label = "Throughput Arrival Timer";
            this.ThroughputArrivalTimerButton.Name = "ThroughputArrivalTimerButton";
            this.ThroughputArrivalTimerButton.ShowImage = true;
            // 
            // BoothTab
            // 
            this.BoothTab.Groups.Add(this.group2);
            this.BoothTab.Groups.Add(this.group3);
            this.BoothTab.Groups.Add(this.group4);
            this.BoothTab.Groups.Add(this.group5);
            this.BoothTab.Groups.Add(this.group6);
            this.BoothTab.Label = "BOOTH";
            this.BoothTab.Name = "BoothTab";
            // 
            // group2
            // 
            this.group2.Items.Add(ImportSplitButton);
            this.group2.Items.Add(ImportPollPadButton);
            this.group2.Label = "Import Data";
            this.group2.Name = "group2";
            // 
            // ImportPollPadButton
            // 
            ImportPollPadButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            ImportPollPadButton.Label = "Import PollPad File(s)";
            ImportPollPadButton.Name = "ImportPollPadButton";
            ImportPollPadButton.OfficeImageId = "GetExternalDataImportClassic";
            ImportPollPadButton.ShowImage = true;
            ImportPollPadButton.SuperTip = "Import PollPad log file(s) in *.txt* and *.csv* format. Create new Worksheet(s) a" +
    "nd populate with the selected file(s).";
            ImportPollPadButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ImportButton_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.ProcessSingleButton);
            this.group3.Items.Add(this.ProcessAllButton);
            this.group3.Label = "Process Data";
            this.group3.Name = "group3";
            // 
            // ProcessSingleButton
            // 
            this.ProcessSingleButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ProcessSingleButton.Label = "Single Data Sheet";
            this.ProcessSingleButton.Name = "ProcessSingleButton";
            this.ProcessSingleButton.OfficeImageId = "QueryCrosstab";
            this.ProcessSingleButton.ShowImage = true;
            this.ProcessSingleButton.SuperTip = "Create a new Worksheet and populate with processed data from the active Worksheet" +
    ".";
            this.ProcessSingleButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ProcessButton_Click);
            // 
            // ProcessAllButton
            // 
            this.ProcessAllButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ProcessAllButton.Label = "All Open Data Sheets";
            this.ProcessAllButton.Name = "ProcessAllButton";
            this.ProcessAllButton.OfficeImageId = "QuerySelectQueryType";
            this.ProcessAllButton.ShowImage = true;
            this.ProcessAllButton.SuperTip = "Create new Worksheets populated with processed data from every applicable Workshe" +
    "et in an open Workbook.";
            this.ProcessAllButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ProcessButton_Click);
            // 
            // group4
            // 
            this.group4.Items.Add(ProcessFolderButton);
            this.group4.Label = "Process Entire Folder";
            this.group4.Name = "group4";
            // 
            // group5
            // 
            this.group5.Items.Add(this.CreateSumStatsButton);
            this.group5.Label = "Analysis";
            this.group5.Name = "group5";
            // 
            // CreateSumStatsButton
            // 
            this.CreateSumStatsButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.CreateSumStatsButton.Label = "Create Summary Statistics";
            this.CreateSumStatsButton.Name = "CreateSumStatsButton";
            this.CreateSumStatsButton.OfficeImageId = "WhatIfAnalysisMenu";
            this.CreateSumStatsButton.ShowImage = true;
            this.CreateSumStatsButton.SuperTip = "Generates summary statistics for the open worksheet. Autodetects the type of data" +
    " inserted. Can only be used on already processed sheets.";
            // 
            // group6
            // 
            this.group6.Items.Add(TimersButton);
            this.group6.Label = "Voting Timers";
            this.group6.Name = "group6";
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.BoothTab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.BoothTab.ResumeLayout(false);
            this.BoothTab.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.group4.ResumeLayout(false);
            this.group4.PerformLayout();
            this.group5.ResumeLayout(false);
            this.group5.PerformLayout();
            this.group6.ResumeLayout(false);
            this.group6.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group4;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group5;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group6;
        private Microsoft.Office.Tools.Ribbon.RibbonTab BoothTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ImportDS200Button;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ImportDICEButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ImportVSAPBMDButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ImportDICXButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ProcessAllButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ProcessSingleButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ProcessDS200Button;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ProcessVSAPBMDButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ProcessDICEButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ProcessDICXButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton CreateSumStatsButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton CheckinArrivalTimerButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton VotingBoothTimerButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BMDTimerButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BallotScanningTimerButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ThroughputArrivalTimerButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton CheckinTimerButton;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
