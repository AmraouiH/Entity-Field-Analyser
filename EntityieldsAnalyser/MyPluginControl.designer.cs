using EntityieldsAnalyser.UI;

namespace EntityieldsAnalyser
{
    partial class MyPluginControl
    {
        /// <summary>
        /// Variable nécessaire au concepteur.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Nettoyage des ressources utilisées.
        /// </summary>
        /// <param name="disposing">true si les ressources managées doivent être supprimées ; sinon, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Code généré par le Concepteur de composants

        /// <summary>
        /// Méthode requise pour la prise en charge du concepteur - ne modifiez pas
        /// le contenu de cette méthode avec l'éditeur de code.
        /// Charts are created in code (MyPluginControl.Insights.cs) and docked in the chart cards below.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.toolStripMenu = new System.Windows.Forms.ToolStrip();
            this.tsbClose = new System.Windows.Forms.ToolStripButton();
            this.tssSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.entityTypeComboBox = new EntityieldsAnalyser.UI.ToolStripFluentComboBox();
            this.toolStripButton = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.AnalyseType = new EntityieldsAnalyser.UI.ToolStripFluentComboBox();
            this.analyseButton = new System.Windows.Forms.ToolStripButton();
            this.buttonExport = new System.Windows.Forms.ToolStripButton();
            this.byButton = new System.Windows.Forms.ToolStripButton();
            this.helpButton = new System.Windows.Forms.ToolStripButton();
            this.buymeacoffeeicon = new System.Windows.Forms.ToolStripButton();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.groupBox1 = new EntityieldsAnalyser.UI.CardBox();
            this.entitiesLayout = new System.Windows.Forms.TableLayoutPanel();
            this.searchEntity = new EntityieldsAnalyser.UI.FluentTextBox();
            this.EntityGridView = new System.Windows.Forms.DataGridView();
            this.rightLayout = new System.Windows.Forms.TableLayoutPanel();
            this.contextHeader = new EntityieldsAnalyser.UI.ContextHeader();
            this.pivotTabs = new EntityieldsAnalyser.UI.PivotTabs();
            this.pagesHost = new System.Windows.Forms.Panel();
            // Welcome
            this.pageWelcome = new System.Windows.Forms.TableLayoutPanel();
            this.welcomeTitle = new System.Windows.Forms.Label();
            this.stepsLayout = new System.Windows.Forms.TableLayoutPanel();
            this.stepLoad = new EntityieldsAnalyser.UI.StepCard();
            this.stepPick = new EntityieldsAnalyser.UI.StepCard();
            this.stepAnalyse = new EntityieldsAnalyser.UI.StepCard();
            this.guideActionPanel = new System.Windows.Forms.FlowLayoutPanel();
            this.guideAction = new EntityieldsAnalyser.UI.FluentButton();
            this.guideHint = new System.Windows.Forms.Label();
            this.featuresTitle = new System.Windows.Forms.Label();
            this.featuresLayout = new System.Windows.Forms.TableLayoutPanel();
            this.featureOverview = new EntityieldsAnalyser.UI.CardBox();
            this.featureOverviewText = new System.Windows.Forms.Label();
            this.featureFields = new EntityieldsAnalyser.UI.CardBox();
            this.featureFieldsText = new System.Windows.Forms.Label();
            this.featureUsage = new EntityieldsAnalyser.UI.CardBox();
            this.featureUsageText = new System.Windows.Forms.Label();
            this.featureCapacity = new EntityieldsAnalyser.UI.CardBox();
            this.featureCapacityText = new System.Windows.Forms.Label();
            // Overview
            this.pageOverview = new System.Windows.Forms.TableLayoutPanel();
            this.kpiLayout = new System.Windows.Forms.TableLayoutPanel();
            this.kpiFields = new EntityieldsAnalyser.UI.KpiTile();
            this.kpiCustom = new EntityieldsAnalyser.UI.KpiTile();
            this.kpiUnmanaged = new EntityieldsAnalyser.UI.KpiTile();
            this.kpiCapacity = new EntityieldsAnalyser.UI.KpiTile();
            this.cardTypes = new EntityieldsAnalyser.UI.CardBox();
            this.cardOrigin = new EntityieldsAnalyser.UI.CardBox();
            this.cardSettings = new EntityieldsAnalyser.UI.CardBox();
            this.cardPublishers = new EntityieldsAnalyser.UI.CardBox();
            this.cardTimeline = new EntityieldsAnalyser.UI.CardBox();
            // Fields
            this.pageFields = new System.Windows.Forms.Panel();
            this.fieldsGroupBox = new EntityieldsAnalyser.UI.CardBox();
            this.tableLayoutPanel4 = new System.Windows.Forms.TableLayoutPanel();
            this.tableLayoutPanel5 = new System.Windows.Forms.TableLayoutPanel();
            this.fieldTypeCombobox = new EntityieldsAnalyser.UI.FluentComboBox();
            this.searchField = new EntityieldsAnalyser.UI.FluentTextBox();
            this.displayAllColumns = new EntityieldsAnalyser.UI.ToggleSwitch();
            this.fieldsFilterRow = new System.Windows.Forms.TableLayoutPanel();
            this.fieldsChips = new EntityieldsAnalyser.UI.ChipBar();
            this.fieldsResultLabel = new System.Windows.Forms.Label();
            this.fieldPropretiesView = new System.Windows.Forms.DataGridView();
            // Data usage
            this.pageUsage = new System.Windows.Forms.Panel();
            this.usageLayout = new System.Windows.Forms.TableLayoutPanel();
            this.usageKpiLayout = new System.Windows.Forms.TableLayoutPanel();
            this.kpiRecords = new EntityieldsAnalyser.UI.KpiTile();
            this.kpiUnused = new EntityieldsAnalyser.UI.KpiTile();
            this.kpiLowUsage = new EntityieldsAnalyser.UI.KpiTile();
            this.kpiAverage = new EntityieldsAnalyser.UI.KpiTile();
            this.cardFillDistribution = new EntityieldsAnalyser.UI.CardBox();
            this.cardFillByType = new EntityieldsAnalyser.UI.CardBox();
            this.cardLeastUsed = new EntityieldsAnalyser.UI.CardBox();
            this.usageEmpty = new EntityieldsAnalyser.UI.CardBox();
            this.usageEmptyLayout = new System.Windows.Forms.TableLayoutPanel();
            this.usageEmptyText = new System.Windows.Forms.Label();
            this.rerunWithUsage = new EntityieldsAnalyser.UI.FluentButton();
            // Capacity
            this.pageCapacity = new System.Windows.Forms.TableLayoutPanel();
            this.cardBudget = new EntityieldsAnalyser.UI.CardBox();
            this.budgetLayout = new System.Windows.Forms.TableLayoutPanel();
            this.budgetHeadline = new System.Windows.Forms.Label();
            this.fieldCalculatorGroupBox = new EntityieldsAnalyser.UI.CardBox();
            this.tableLayoutPanel7 = new System.Windows.Forms.TableLayoutPanel();
            this.label5 = new System.Windows.Forms.Label();
            this.tableLayoutPanel6 = new System.Windows.Forms.TableLayoutPanel();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.textBox1 = new EntityieldsAnalyser.UI.FluentTextBox();
            this.textBox2 = new EntityieldsAnalyser.UI.FluentTextBox();
            this.textBox3 = new EntityieldsAnalyser.UI.FluentTextBox();
            this.calculatorResult = new System.Windows.Forms.Label();
            this.MoreDetailsLink = new System.Windows.Forms.Label();
            this.cardColumnCost = new EntityieldsAnalyser.UI.CardBox();
            this.toolStripMenu.SuspendLayout();
            this.tableLayoutPanel1.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.entitiesLayout.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.EntityGridView)).BeginInit();
            this.rightLayout.SuspendLayout();
            this.pagesHost.SuspendLayout();
            this.pageWelcome.SuspendLayout();
            this.stepsLayout.SuspendLayout();
            this.guideActionPanel.SuspendLayout();
            this.featuresLayout.SuspendLayout();
            this.pageOverview.SuspendLayout();
            this.kpiLayout.SuspendLayout();
            this.pageFields.SuspendLayout();
            this.fieldsGroupBox.SuspendLayout();
            this.tableLayoutPanel4.SuspendLayout();
            this.tableLayoutPanel5.SuspendLayout();
            this.fieldsFilterRow.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.fieldPropretiesView)).BeginInit();
            this.pageUsage.SuspendLayout();
            this.usageLayout.SuspendLayout();
            this.usageKpiLayout.SuspendLayout();
            this.usageEmpty.SuspendLayout();
            this.usageEmptyLayout.SuspendLayout();
            this.pageCapacity.SuspendLayout();
            this.cardBudget.SuspendLayout();
            this.budgetLayout.SuspendLayout();
            this.fieldCalculatorGroupBox.SuspendLayout();
            this.tableLayoutPanel7.SuspendLayout();
            this.tableLayoutPanel6.SuspendLayout();
            this.SuspendLayout();
            //
            // toolStripMenu
            //
            this.toolStripMenu.AutoSize = false;
            this.toolStripMenu.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden;
            this.toolStripMenu.ImageScalingSize = new System.Drawing.Size(16, 16);
            this.toolStripMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tsbClose,
            this.tssSeparator1,
            this.entityTypeComboBox,
            this.toolStripButton,
            this.toolStripSeparator1,
            this.AnalyseType,
            this.analyseButton,
            this.buttonExport,
            this.byButton,
            this.helpButton,
            this.buymeacoffeeicon});
            this.toolStripMenu.Location = new System.Drawing.Point(0, 0);
            this.toolStripMenu.Name = "toolStripMenu";
            this.toolStripMenu.Padding = new System.Windows.Forms.Padding(10, 6, 10, 6);
            this.toolStripMenu.Size = new System.Drawing.Size(1913, 48);
            this.toolStripMenu.TabIndex = 4;
            //
            // tsbClose
            //
            this.tsbClose.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.tsbClose.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
            this.tsbClose.Margin = new System.Windows.Forms.Padding(0, 0, 2, 0);
            this.tsbClose.Name = "tsbClose";
            this.tsbClose.Overflow = System.Windows.Forms.ToolStripItemOverflow.Never;
            this.tsbClose.Padding = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.tsbClose.Size = new System.Drawing.Size(32, 36);
            this.tsbClose.ToolTipText = "Close Tool";
            this.tsbClose.Click += new System.EventHandler(this.tsbClose_Click);
            //
            // tssSeparator1
            //
            this.tssSeparator1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.tssSeparator1.Name = "tssSeparator1";
            this.tssSeparator1.Size = new System.Drawing.Size(6, 36);
            //
            // entityTypeComboBox
            //
            this.entityTypeComboBox.DropDownWidth = 160;
            this.entityTypeComboBox.Items.AddRange(new object[] {
            "All Entities",
            "Custom Entities",
            "CRM Entities",
            "System Entities"});
            this.entityTypeComboBox.Margin = new System.Windows.Forms.Padding(6, 2, 4, 2);
            this.entityTypeComboBox.Name = "entityTypeComboBox";
            this.entityTypeComboBox.Size = new System.Drawing.Size(150, 32);
            this.entityTypeComboBox.ToolTipText = "Select the Type Of Entities You Want to Analyse";
            //
            // toolStripButton
            //
            this.toolStripButton.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
            this.toolStripButton.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.toolStripButton.Name = "toolStripButton";
            this.toolStripButton.Padding = new System.Windows.Forms.Padding(8, 0, 8, 0);
            this.toolStripButton.Size = new System.Drawing.Size(110, 36);
            this.toolStripButton.Text = "Get Entities";
            this.toolStripButton.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.toolStripButton.ToolTipText = "Get the Entities";
            this.toolStripButton.Click += new System.EventHandler(this.getEntitiesButton_Click);
            //
            // toolStripSeparator1
            //
            this.toolStripSeparator1.Margin = new System.Windows.Forms.Padding(8, 0, 8, 0);
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(6, 36);
            //
            // AnalyseType
            //
            this.AnalyseType.DropDownWidth = 200;
            this.AnalyseType.Items.AddRange(new object[] {
            "Metadata Only",
            "Metadata + Data usage"});
            this.AnalyseType.Margin = new System.Windows.Forms.Padding(0, 2, 6, 2);
            this.AnalyseType.Name = "AnalyseType";
            this.AnalyseType.Size = new System.Drawing.Size(190, 32);
            this.AnalyseType.ToolTipText = "The second option takes much time depends on your entity data size";
            this.AnalyseType.SelectedIndexChanged += new System.EventHandler(this.AnalyseType_SelectedIndexChanged);
            //
            // analyseButton
            //
            this.analyseButton.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
            this.analyseButton.Margin = new System.Windows.Forms.Padding(2, 0, 4, 0);
            this.analyseButton.Name = "analyseButton";
            this.analyseButton.Padding = new System.Windows.Forms.Padding(10, 0, 12, 0);
            this.analyseButton.Size = new System.Drawing.Size(96, 36);
            this.analyseButton.Tag = "primary";
            this.analyseButton.Text = "Analyse";
            this.analyseButton.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.analyseButton.ToolTipText = "Run Analyse to get Results";
            this.analyseButton.Click += new System.EventHandler(this.analyseButton_Click);
            //
            // buttonExport
            //
            this.buttonExport.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
            this.buttonExport.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.buttonExport.Name = "buttonExport";
            this.buttonExport.Padding = new System.Windows.Forms.Padding(8, 0, 8, 0);
            this.buttonExport.Size = new System.Drawing.Size(130, 36);
            this.buttonExport.Text = "Export to Excel";
            this.buttonExport.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.buttonExport.ToolTipText = "Export the Results to Excel File";
            this.buttonExport.Click += new System.EventHandler(this.buttonExport_Click);
            //
            // byButton
            //
            this.byButton.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.byButton.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
            this.byButton.Margin = new System.Windows.Forms.Padding(2, 0, 0, 0);
            this.byButton.Name = "byButton";
            this.byButton.Padding = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.byButton.Size = new System.Drawing.Size(150, 36);
            this.byButton.Tag = "link";
            this.byButton.Text = "Hamza AMRAOUI";
            this.byButton.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.byButton.ToolTipText = "Developed by Hamza AMRAOUI (LinkedIn)";
            this.byButton.Click += new System.EventHandler(this.byButton_Click);
            //
            // helpButton
            //
            this.helpButton.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.helpButton.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
            this.helpButton.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.helpButton.Name = "helpButton";
            this.helpButton.Padding = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.helpButton.Size = new System.Drawing.Size(70, 36);
            this.helpButton.Tag = "link";
            this.helpButton.Text = "Help";
            this.helpButton.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.helpButton.ToolTipText = "How to use the tool, report an issue";
            this.helpButton.Click += new System.EventHandler(this.toolStripButton1_Click);
            //
            // buymeacoffeeicon
            //
            this.buymeacoffeeicon.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.buymeacoffeeicon.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
            this.buymeacoffeeicon.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.buymeacoffeeicon.Name = "buymeacoffeeicon";
            this.buymeacoffeeicon.Padding = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.buymeacoffeeicon.Size = new System.Drawing.Size(90, 36);
            this.buymeacoffeeicon.Tag = "link";
            this.buymeacoffeeicon.Text = "Donate";
            this.buymeacoffeeicon.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.buymeacoffeeicon.ToolTipText = "Support the tool (PayPal)";
            this.buymeacoffeeicon.Click += new System.EventHandler(this.buymeacoffeeiconClick);
            //
            // tableLayoutPanel1
            //
            this.tableLayoutPanel1.ColumnCount = 2;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 22F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 78F));
            this.tableLayoutPanel1.Controls.Add(this.groupBox1, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.rightLayout, 1, 0);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Margin = new System.Windows.Forms.Padding(0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.Padding = new System.Windows.Forms.Padding(8);
            this.tableLayoutPanel1.RowCount = 1;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.TabIndex = 15;
            //
            // groupBox1
            //
            this.groupBox1.Controls.Add(this.entitiesLayout);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Subtitle = "Load, then pick one to analyse";
            this.groupBox1.TabIndex = 21;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Entities";
            //
            // entitiesLayout
            //
            this.entitiesLayout.ColumnCount = 1;
            this.entitiesLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.entitiesLayout.Controls.Add(this.searchEntity, 0, 0);
            this.entitiesLayout.Controls.Add(this.EntityGridView, 0, 1);
            this.entitiesLayout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.entitiesLayout.Margin = new System.Windows.Forms.Padding(0);
            this.entitiesLayout.Name = "entitiesLayout";
            this.entitiesLayout.RowCount = 2;
            this.entitiesLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 42F));
            this.entitiesLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            //
            // searchEntity
            //
            this.searchEntity.Dock = System.Windows.Forms.DockStyle.Top;
            this.searchEntity.Glyph = EntityieldsAnalyser.UI.Theme.Glyph.Search;
            this.searchEntity.Margin = new System.Windows.Forms.Padding(0, 0, 0, 10);
            this.searchEntity.Name = "searchEntity";
            this.searchEntity.Placeholder = "Search by display or schema name";
            this.searchEntity.Size = new System.Drawing.Size(300, 32);
            this.searchEntity.TabIndex = 0;
            this.searchEntity.Click += new System.EventHandler(this.textBox1_Click);
            this.searchEntity.TextChanged += new System.EventHandler(this.searchEntity_TextChaneged);
            //
            // EntityGridView
            //
            this.EntityGridView.AllowUserToAddRows = false;
            this.EntityGridView.AllowUserToDeleteRows = false;
            this.EntityGridView.AllowUserToResizeRows = false;
            this.EntityGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            this.EntityGridView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.EntityGridView.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically;
            this.EntityGridView.Margin = new System.Windows.Forms.Padding(0);
            this.EntityGridView.MultiSelect = false;
            this.EntityGridView.Name = "EntityGridView";
            this.EntityGridView.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            this.EntityGridView.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.EntityGridView.ShowEditingIcon = false;
            this.EntityGridView.TabIndex = 3;
            this.EntityGridView.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.EntityGridView_CellContentClick_1);
            this.EntityGridView.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.EntityGridView_CellDoubleClick);
            //
            // rightLayout
            //
            this.rightLayout.ColumnCount = 1;
            this.rightLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.rightLayout.Controls.Add(this.contextHeader, 0, 0);
            this.rightLayout.Controls.Add(this.pivotTabs, 0, 1);
            this.rightLayout.Controls.Add(this.pagesHost, 0, 2);
            this.rightLayout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.rightLayout.Margin = new System.Windows.Forms.Padding(0);
            this.rightLayout.Name = "rightLayout";
            this.rightLayout.RowCount = 3;
            this.rightLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 68F));
            this.rightLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 46F));
            this.rightLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            //
            // contextHeader
            //
            this.contextHeader.Dock = System.Windows.Forms.DockStyle.Fill;
            this.contextHeader.Margin = new System.Windows.Forms.Padding(8, 4, 8, 0);
            this.contextHeader.Name = "contextHeader";
            //
            // pivotTabs
            //
            this.pivotTabs.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pivotTabs.Margin = new System.Windows.Forms.Padding(8, 0, 8, 4);
            this.pivotTabs.Name = "pivotTabs";
            this.pivotTabs.TabIndex = 1;
            this.pivotTabs.SelectedIndexChanged += new System.EventHandler(this.pivotTabs_SelectedIndexChanged);
            //
            // pagesHost
            //
            this.pagesHost.Controls.Add(this.pageWelcome);
            this.pagesHost.Controls.Add(this.pageOverview);
            this.pagesHost.Controls.Add(this.pageFields);
            this.pagesHost.Controls.Add(this.pageUsage);
            this.pagesHost.Controls.Add(this.pageCapacity);
            this.pagesHost.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pagesHost.Margin = new System.Windows.Forms.Padding(0);
            this.pagesHost.Name = "pagesHost";
            //
            // pageWelcome
            //
            this.pageWelcome.ColumnCount = 1;
            this.pageWelcome.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.pageWelcome.Controls.Add(this.welcomeTitle, 0, 0);
            this.pageWelcome.Controls.Add(this.stepsLayout, 0, 1);
            this.pageWelcome.Controls.Add(this.guideActionPanel, 0, 2);
            this.pageWelcome.Controls.Add(this.featuresTitle, 0, 3);
            this.pageWelcome.Controls.Add(this.featuresLayout, 0, 4);
            this.pageWelcome.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pageWelcome.Name = "pageWelcome";
            this.pageWelcome.RowCount = 6;
            this.pageWelcome.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 40F));
            this.pageWelcome.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 200F));
            this.pageWelcome.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 64F));
            this.pageWelcome.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 40F));
            this.pageWelcome.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 140F));
            this.pageWelcome.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            //
            // welcomeTitle
            //
            this.welcomeTitle.Dock = System.Windows.Forms.DockStyle.Fill;
            this.welcomeTitle.Margin = new System.Windows.Forms.Padding(8, 0, 8, 0);
            this.welcomeTitle.Name = "welcomeTitle";
            this.welcomeTitle.Text = "Get started in 3 steps";
            this.welcomeTitle.TextAlign = System.Drawing.ContentAlignment.BottomLeft;
            //
            // stepsLayout
            //
            this.stepsLayout.ColumnCount = 3;
            this.stepsLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 33.33F));
            this.stepsLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 33.33F));
            this.stepsLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 33.34F));
            this.stepsLayout.Controls.Add(this.stepLoad, 0, 0);
            this.stepsLayout.Controls.Add(this.stepPick, 1, 0);
            this.stepsLayout.Controls.Add(this.stepAnalyse, 2, 0);
            this.stepsLayout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.stepsLayout.Margin = new System.Windows.Forms.Padding(0);
            this.stepsLayout.Name = "stepsLayout";
            this.stepsLayout.RowCount = 1;
            this.stepsLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            //
            // stepLoad
            //
            this.stepLoad.Description = "Choose which entities to list in the toolbar filter (all, custom, CRM or system), then click “Get Entities”.";
            this.stepLoad.Dock = System.Windows.Forms.DockStyle.Fill;
            this.stepLoad.Name = "stepLoad";
            this.stepLoad.Number = 1;
            this.stepLoad.Text = "Load the entities";
            //
            // stepPick
            //
            this.stepPick.Description = "Click an entity in the list on the left. Use the search box to find it by display or schema name.";
            this.stepPick.Dock = System.Windows.Forms.DockStyle.Fill;
            this.stepPick.Name = "stepPick";
            this.stepPick.Number = 2;
            this.stepPick.Text = "Choose an entity";
            //
            // stepAnalyse
            //
            this.stepAnalyse.Description = "“Metadata Only” is instant. “Metadata + Data usage” also reads the records to measure how often each field is filled.";
            this.stepAnalyse.Dock = System.Windows.Forms.DockStyle.Fill;
            this.stepAnalyse.Name = "stepAnalyse";
            this.stepAnalyse.Number = 3;
            this.stepAnalyse.Text = "Run the analysis";
            //
            // guideActionPanel
            //
            this.guideActionPanel.Controls.Add(this.guideAction);
            this.guideActionPanel.Controls.Add(this.guideHint);
            this.guideActionPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.guideActionPanel.Margin = new System.Windows.Forms.Padding(6, 6, 6, 0);
            this.guideActionPanel.Name = "guideActionPanel";
            this.guideActionPanel.WrapContents = false;
            //
            // guideAction
            //
            this.guideAction.Glyph = EntityieldsAnalyser.UI.Theme.Glyph.List;
            this.guideAction.Margin = new System.Windows.Forms.Padding(0, 6, 14, 0);
            this.guideAction.Name = "guideAction";
            this.guideAction.Size = new System.Drawing.Size(170, 36);
            this.guideAction.TabIndex = 0;
            this.guideAction.Text = "Get Entities";
            this.guideAction.Click += new System.EventHandler(this.guideAction_Click);
            //
            // guideHint
            //
            this.guideHint.AutoSize = true;
            this.guideHint.Margin = new System.Windows.Forms.Padding(0, 15, 0, 0);
            this.guideHint.Name = "guideHint";
            //
            // featuresTitle
            //
            this.featuresTitle.Dock = System.Windows.Forms.DockStyle.Fill;
            this.featuresTitle.Margin = new System.Windows.Forms.Padding(8, 0, 8, 0);
            this.featuresTitle.Name = "featuresTitle";
            this.featuresTitle.Text = "What you will find after the analysis";
            this.featuresTitle.TextAlign = System.Drawing.ContentAlignment.BottomLeft;
            //
            // featuresLayout
            //
            this.featuresLayout.ColumnCount = 4;
            this.featuresLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25F));
            this.featuresLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25F));
            this.featuresLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25F));
            this.featuresLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25F));
            this.featuresLayout.Controls.Add(this.featureOverview, 0, 0);
            this.featuresLayout.Controls.Add(this.featureFields, 1, 0);
            this.featuresLayout.Controls.Add(this.featureUsage, 2, 0);
            this.featuresLayout.Controls.Add(this.featureCapacity, 3, 0);
            this.featuresLayout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.featuresLayout.Margin = new System.Windows.Forms.Padding(0);
            this.featuresLayout.Name = "featuresLayout";
            this.featuresLayout.RowCount = 1;
            this.featuresLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            //
            // featureOverview
            //
            this.featureOverview.Controls.Add(this.featureOverviewText);
            this.featureOverview.Dock = System.Windows.Forms.DockStyle.Fill;
            this.featureOverview.Name = "featureOverview";
            this.featureOverview.TabStop = false;
            this.featureOverview.Text = "Overview";
            this.featureOverviewText.Dock = System.Windows.Forms.DockStyle.Fill;
            this.featureOverviewText.Name = "featureOverviewText";
            this.featureOverviewText.Text = "Key figures, field types, origin (standard / custom, managed / unmanaged), publishers and creation timeline.";
            //
            // featureFields
            //
            this.featureFields.Controls.Add(this.featureFieldsText);
            this.featureFields.Dock = System.Windows.Forms.DockStyle.Fill;
            this.featureFields.Name = "featureFields";
            this.featureFields.TabStop = false;
            this.featureFields.Text = "Fields";
            this.featureFieldsText.Dock = System.Windows.Forms.DockStyle.Fill;
            this.featureFieldsText.Name = "featureFieldsText";
            this.featureFieldsText.Text = "Every field with its metadata. Filter by type, origin or usage. Click any chart to open the matching fields here.";
            //
            // featureUsage
            //
            this.featureUsage.Controls.Add(this.featureUsageText);
            this.featureUsage.Dock = System.Windows.Forms.DockStyle.Fill;
            this.featureUsage.Name = "featureUsage";
            this.featureUsage.TabStop = false;
            this.featureUsage.Text = "Data usage";
            this.featureUsageText.Dock = System.Windows.Forms.DockStyle.Fill;
            this.featureUsageText.Name = "featureUsageText";
            this.featureUsageText.Text = "How often each field is filled, never used fields and clean-up candidates (requires “Metadata + Data usage”).";
            //
            // featureCapacity
            //
            this.featureCapacity.Controls.Add(this.featureCapacityText);
            this.featureCapacity.Dock = System.Windows.Forms.DockStyle.Fill;
            this.featureCapacity.Name = "featureCapacity";
            this.featureCapacity.TabStop = false;
            this.featureCapacity.Text = "Capacity";
            this.featureCapacityText.Dock = System.Windows.Forms.DockStyle.Fill;
            this.featureCapacityText.Name = "featureCapacityText";
            this.featureCapacityText.Text = "The physical column budget (1,024 per entity), what consumes it, and a calculator to plan new fields.";
            //
            // pageOverview
            //
            this.pageOverview.ColumnCount = 3;
            this.pageOverview.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 38F));
            this.pageOverview.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 31F));
            this.pageOverview.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 31F));
            this.pageOverview.Controls.Add(this.kpiLayout, 0, 0);
            this.pageOverview.SetColumnSpan(this.kpiLayout, 3);
            this.pageOverview.Controls.Add(this.cardTypes, 0, 1);
            this.pageOverview.Controls.Add(this.cardOrigin, 1, 1);
            this.pageOverview.Controls.Add(this.cardSettings, 2, 1);
            this.pageOverview.Controls.Add(this.cardPublishers, 0, 2);
            this.pageOverview.Controls.Add(this.cardTimeline, 1, 2);
            this.pageOverview.SetColumnSpan(this.cardTimeline, 2);
            this.pageOverview.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pageOverview.Name = "pageOverview";
            this.pageOverview.RowCount = 3;
            this.pageOverview.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 96F));
            this.pageOverview.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 52F));
            this.pageOverview.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 48F));
            this.pageOverview.Visible = false;
            //
            // kpiLayout
            //
            this.kpiLayout.ColumnCount = 4;
            this.kpiLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25F));
            this.kpiLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25F));
            this.kpiLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25F));
            this.kpiLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25F));
            this.kpiLayout.Controls.Add(this.kpiFields, 0, 0);
            this.kpiLayout.Controls.Add(this.kpiCustom, 1, 0);
            this.kpiLayout.Controls.Add(this.kpiUnmanaged, 2, 0);
            this.kpiLayout.Controls.Add(this.kpiCapacity, 3, 0);
            this.kpiLayout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.kpiLayout.Margin = new System.Windows.Forms.Padding(0);
            this.kpiLayout.Name = "kpiLayout";
            this.kpiLayout.RowCount = 1;
            this.kpiLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            //
            // kpiFields
            //
            this.kpiFields.Clickable = true;
            this.kpiFields.Dock = System.Windows.Forms.DockStyle.Fill;
            this.kpiFields.Glyph = EntityieldsAnalyser.UI.Theme.Glyph.List;
            this.kpiFields.Name = "kpiFields";
            this.kpiFields.Title = "Fields";
            this.kpiFields.Click += new System.EventHandler(this.kpiFields_Click);
            //
            // kpiCustom
            //
            this.kpiCustom.Clickable = true;
            this.kpiCustom.Dock = System.Windows.Forms.DockStyle.Fill;
            this.kpiCustom.Glyph = EntityieldsAnalyser.UI.Theme.Glyph.Personalize;
            this.kpiCustom.Name = "kpiCustom";
            this.kpiCustom.Title = "Custom fields";
            this.kpiCustom.Click += new System.EventHandler(this.kpiCustom_Click);
            //
            // kpiUnmanaged
            //
            this.kpiUnmanaged.Clickable = true;
            this.kpiUnmanaged.Dock = System.Windows.Forms.DockStyle.Fill;
            this.kpiUnmanaged.Glyph = EntityieldsAnalyser.UI.Theme.Glyph.Unlock;
            this.kpiUnmanaged.Name = "kpiUnmanaged";
            this.kpiUnmanaged.Title = "Unmanaged fields";
            this.kpiUnmanaged.Click += new System.EventHandler(this.kpiUnmanaged_Click);
            //
            // kpiCapacity
            //
            this.kpiCapacity.Clickable = true;
            this.kpiCapacity.Dock = System.Windows.Forms.DockStyle.Fill;
            this.kpiCapacity.Glyph = EntityieldsAnalyser.UI.Theme.Glyph.Package;
            this.kpiCapacity.Name = "kpiCapacity";
            this.kpiCapacity.Progress = 0F;
            this.kpiCapacity.Title = "Column capacity used";
            this.kpiCapacity.Click += new System.EventHandler(this.kpiCapacity_Click);
            //
            // overview cards
            //
            this.cardTypes.Dock = System.Windows.Forms.DockStyle.Fill;
            this.cardTypes.Name = "cardTypes";
            this.cardTypes.Subtitle = "click a bar to list them";
            this.cardTypes.TabStop = false;
            this.cardTypes.Text = "Field types";
            this.cardOrigin.Dock = System.Windows.Forms.DockStyle.Fill;
            this.cardOrigin.Name = "cardOrigin";
            this.cardOrigin.TabStop = false;
            this.cardOrigin.Text = "Origin";
            this.cardSettings.Dock = System.Windows.Forms.DockStyle.Fill;
            this.cardSettings.Name = "cardSettings";
            this.cardSettings.Subtitle = "% of fields";
            this.cardSettings.TabStop = false;
            this.cardSettings.Text = "Field settings";
            this.cardPublishers.Dock = System.Windows.Forms.DockStyle.Fill;
            this.cardPublishers.Name = "cardPublishers";
            this.cardPublishers.Subtitle = "custom fields by prefix";
            this.cardPublishers.TabStop = false;
            this.cardPublishers.Text = "Publishers";
            this.cardTimeline.Dock = System.Windows.Forms.DockStyle.Fill;
            this.cardTimeline.Name = "cardTimeline";
            this.cardTimeline.TabStop = false;
            this.cardTimeline.Text = "Fields created over time";
            //
            // pageFields
            //
            this.pageFields.Controls.Add(this.fieldsGroupBox);
            this.pageFields.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pageFields.Name = "pageFields";
            this.pageFields.Visible = false;
            //
            // fieldsGroupBox
            //
            this.fieldsGroupBox.Controls.Add(this.tableLayoutPanel4);
            this.fieldsGroupBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.fieldsGroupBox.Name = "fieldsGroupBox";
            this.fieldsGroupBox.TabIndex = 25;
            this.fieldsGroupBox.TabStop = false;
            this.fieldsGroupBox.Text = "Entity fields";
            //
            // tableLayoutPanel4
            //
            this.tableLayoutPanel4.ColumnCount = 1;
            this.tableLayoutPanel4.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel4.Controls.Add(this.tableLayoutPanel5, 0, 0);
            this.tableLayoutPanel4.Controls.Add(this.fieldsFilterRow, 0, 1);
            this.tableLayoutPanel4.Controls.Add(this.fieldPropretiesView, 0, 2);
            this.tableLayoutPanel4.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel4.Margin = new System.Windows.Forms.Padding(0);
            this.tableLayoutPanel4.Name = "tableLayoutPanel4";
            this.tableLayoutPanel4.RowCount = 3;
            this.tableLayoutPanel4.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 42F));
            this.tableLayoutPanel4.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 44F));
            this.tableLayoutPanel4.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            //
            // tableLayoutPanel5
            //
            this.tableLayoutPanel5.ColumnCount = 4;
            this.tableLayoutPanel5.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 230F));
            this.tableLayoutPanel5.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 320F));
            this.tableLayoutPanel5.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel5.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.AutoSize));
            this.tableLayoutPanel5.Controls.Add(this.fieldTypeCombobox, 0, 0);
            this.tableLayoutPanel5.Controls.Add(this.searchField, 1, 0);
            this.tableLayoutPanel5.Controls.Add(this.displayAllColumns, 3, 0);
            this.tableLayoutPanel5.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel5.Margin = new System.Windows.Forms.Padding(0);
            this.tableLayoutPanel5.Name = "tableLayoutPanel5";
            this.tableLayoutPanel5.RowCount = 1;
            this.tableLayoutPanel5.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            //
            // fieldTypeCombobox
            //
            this.fieldTypeCombobox.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.fieldTypeCombobox.Margin = new System.Windows.Forms.Padding(0, 0, 8, 0);
            this.fieldTypeCombobox.Name = "fieldTypeCombobox";
            this.fieldTypeCombobox.TabIndex = 5;
            this.fieldTypeCombobox.SelectedIndexChanged += new System.EventHandler(this.fieldTypeCombobox_SelectedIndexChanged);
            //
            // searchField
            //
            this.searchField.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.searchField.Glyph = EntityieldsAnalyser.UI.Theme.Glyph.Search;
            this.searchField.Margin = new System.Windows.Forms.Padding(0, 0, 16, 0);
            this.searchField.Name = "searchField";
            this.searchField.Placeholder = "Search fields by display or schema name";
            this.searchField.Size = new System.Drawing.Size(304, 32);
            this.searchField.TabIndex = 11;
            this.searchField.TextChanged += new System.EventHandler(this.searchField_TextChaneged);
            //
            // displayAllColumns
            //
            this.displayAllColumns.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.displayAllColumns.Margin = new System.Windows.Forms.Padding(0);
            this.displayAllColumns.Name = "displayAllColumns";
            this.displayAllColumns.TabIndex = 10;
            this.displayAllColumns.Text = "All metadata columns";
            this.displayAllColumns.CheckedChanged += new System.EventHandler(this.displayAllColumns_CheckedChanged);
            //
            // fieldsFilterRow
            //
            this.fieldsFilterRow.ColumnCount = 2;
            this.fieldsFilterRow.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.fieldsFilterRow.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.AutoSize));
            this.fieldsFilterRow.Controls.Add(this.fieldsChips, 0, 0);
            this.fieldsFilterRow.Controls.Add(this.fieldsResultLabel, 1, 0);
            this.fieldsFilterRow.Dock = System.Windows.Forms.DockStyle.Fill;
            this.fieldsFilterRow.Margin = new System.Windows.Forms.Padding(0, 0, 0, 6);
            this.fieldsFilterRow.Name = "fieldsFilterRow";
            this.fieldsFilterRow.RowCount = 1;
            this.fieldsFilterRow.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            //
            // fieldsChips
            //
            this.fieldsChips.Dock = System.Windows.Forms.DockStyle.Fill;
            this.fieldsChips.Margin = new System.Windows.Forms.Padding(0);
            this.fieldsChips.Name = "fieldsChips";
            this.fieldsChips.SelectedIndexChanged += new System.EventHandler(this.fieldsChips_SelectedIndexChanged);
            this.fieldsChips.Dismissed += new System.EventHandler(this.fieldsChips_Dismissed);
            //
            // fieldsResultLabel
            //
            this.fieldsResultLabel.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.fieldsResultLabel.AutoSize = true;
            this.fieldsResultLabel.Margin = new System.Windows.Forms.Padding(8, 0, 0, 0);
            this.fieldsResultLabel.Name = "fieldsResultLabel";
            //
            // fieldPropretiesView
            //
            this.fieldPropretiesView.AllowUserToAddRows = false;
            this.fieldPropretiesView.AllowUserToDeleteRows = false;
            this.fieldPropretiesView.AllowUserToResizeRows = false;
            this.fieldPropretiesView.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.fieldPropretiesView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            this.fieldPropretiesView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.fieldPropretiesView.Margin = new System.Windows.Forms.Padding(0);
            this.fieldPropretiesView.Name = "fieldPropretiesView";
            this.fieldPropretiesView.ReadOnly = true;
            this.fieldPropretiesView.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.fieldPropretiesView.TabIndex = 3;
            //
            // pageUsage
            //
            this.pageUsage.Controls.Add(this.usageLayout);
            this.pageUsage.Controls.Add(this.usageEmpty);
            this.pageUsage.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pageUsage.Name = "pageUsage";
            this.pageUsage.Visible = false;
            //
            // usageLayout
            //
            this.usageLayout.ColumnCount = 2;
            this.usageLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.usageLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.usageLayout.Controls.Add(this.usageKpiLayout, 0, 0);
            this.usageLayout.SetColumnSpan(this.usageKpiLayout, 2);
            this.usageLayout.Controls.Add(this.cardFillDistribution, 0, 1);
            this.usageLayout.Controls.Add(this.cardFillByType, 1, 1);
            this.usageLayout.Controls.Add(this.cardLeastUsed, 0, 2);
            this.usageLayout.SetColumnSpan(this.cardLeastUsed, 2);
            this.usageLayout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.usageLayout.Name = "usageLayout";
            this.usageLayout.RowCount = 3;
            this.usageLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 96F));
            this.usageLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.usageLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            //
            // usageKpiLayout
            //
            this.usageKpiLayout.ColumnCount = 4;
            this.usageKpiLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25F));
            this.usageKpiLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25F));
            this.usageKpiLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25F));
            this.usageKpiLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25F));
            this.usageKpiLayout.Controls.Add(this.kpiRecords, 0, 0);
            this.usageKpiLayout.Controls.Add(this.kpiUnused, 1, 0);
            this.usageKpiLayout.Controls.Add(this.kpiLowUsage, 2, 0);
            this.usageKpiLayout.Controls.Add(this.kpiAverage, 3, 0);
            this.usageKpiLayout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.usageKpiLayout.Margin = new System.Windows.Forms.Padding(0);
            this.usageKpiLayout.Name = "usageKpiLayout";
            this.usageKpiLayout.RowCount = 1;
            this.usageKpiLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            //
            // usage KPI tiles
            //
            this.kpiRecords.Dock = System.Windows.Forms.DockStyle.Fill;
            this.kpiRecords.Glyph = EntityieldsAnalyser.UI.Theme.Glyph.Database;
            this.kpiRecords.Name = "kpiRecords";
            this.kpiRecords.Title = "Records analysed";
            this.kpiUnused.Clickable = true;
            this.kpiUnused.Dock = System.Windows.Forms.DockStyle.Fill;
            this.kpiUnused.Glyph = EntityieldsAnalyser.UI.Theme.Glyph.Warning;
            this.kpiUnused.Name = "kpiUnused";
            this.kpiUnused.Title = "Never used fields";
            this.kpiUnused.Click += new System.EventHandler(this.kpiUnused_Click);
            this.kpiLowUsage.Clickable = true;
            this.kpiLowUsage.Dock = System.Windows.Forms.DockStyle.Fill;
            this.kpiLowUsage.Glyph = EntityieldsAnalyser.UI.Theme.Glyph.Chart;
            this.kpiLowUsage.Name = "kpiLowUsage";
            this.kpiLowUsage.Title = "Rarely used (< 5%)";
            this.kpiLowUsage.Click += new System.EventHandler(this.kpiLowUsage_Click);
            this.kpiAverage.Dock = System.Windows.Forms.DockStyle.Fill;
            this.kpiAverage.Glyph = EntityieldsAnalyser.UI.Theme.Glyph.Completed;
            this.kpiAverage.Name = "kpiAverage";
            this.kpiAverage.Progress = 0F;
            this.kpiAverage.Title = "Average fill rate";
            //
            // usage cards
            //
            this.cardFillDistribution.Dock = System.Windows.Forms.DockStyle.Fill;
            this.cardFillDistribution.Name = "cardFillDistribution";
            this.cardFillDistribution.Subtitle = "number of fields per fill rate";
            this.cardFillDistribution.TabStop = false;
            this.cardFillDistribution.Text = "Fill rate distribution";
            this.cardFillByType.Dock = System.Windows.Forms.DockStyle.Fill;
            this.cardFillByType.Name = "cardFillByType";
            this.cardFillByType.Subtitle = "average % of records filled";
            this.cardFillByType.TabStop = false;
            this.cardFillByType.Text = "Fill rate by type";
            this.cardLeastUsed.Dock = System.Windows.Forms.DockStyle.Fill;
            this.cardLeastUsed.Name = "cardLeastUsed";
            this.cardLeastUsed.Subtitle = "custom fields with the lowest fill rate: clean-up candidates";
            this.cardLeastUsed.TabStop = false;
            this.cardLeastUsed.Text = "Least used custom fields";
            //
            // usageEmpty
            //
            this.usageEmpty.Controls.Add(this.usageEmptyLayout);
            this.usageEmpty.Dock = System.Windows.Forms.DockStyle.Top;
            this.usageEmpty.Name = "usageEmpty";
            this.usageEmpty.Size = new System.Drawing.Size(900, 170);
            this.usageEmpty.TabStop = false;
            this.usageEmpty.Text = "Data usage was not analysed";
            this.usageEmptyLayout.ColumnCount = 1;
            this.usageEmptyLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.usageEmptyLayout.Controls.Add(this.usageEmptyText, 0, 0);
            this.usageEmptyLayout.Controls.Add(this.rerunWithUsage, 0, 1);
            this.usageEmptyLayout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.usageEmptyLayout.Margin = new System.Windows.Forms.Padding(0);
            this.usageEmptyLayout.Name = "usageEmptyLayout";
            this.usageEmptyLayout.RowCount = 2;
            this.usageEmptyLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 52F));
            this.usageEmptyLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.usageEmptyText.Dock = System.Windows.Forms.DockStyle.Fill;
            this.usageEmptyText.Margin = new System.Windows.Forms.Padding(0);
            this.usageEmptyText.Name = "usageEmptyText";
            this.usageEmptyText.Text = "This analysis ran in “Metadata Only” mode. Re-run it with “Metadata + Data usage” to measure how often each field is filled and find fields that are never used. It reads every record of the entity, so it can take a while on large tables.";
            this.rerunWithUsage.Glyph = EntityieldsAnalyser.UI.Theme.Glyph.Play;
            this.rerunWithUsage.Margin = new System.Windows.Forms.Padding(0, 4, 0, 0);
            this.rerunWithUsage.Name = "rerunWithUsage";
            this.rerunWithUsage.Size = new System.Drawing.Size(240, 34);
            this.rerunWithUsage.Text = "Analyse with data usage";
            this.rerunWithUsage.Click += new System.EventHandler(this.rerunWithUsage_Click);
            //
            // pageCapacity
            //
            this.pageCapacity.ColumnCount = 1;
            this.pageCapacity.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.pageCapacity.Controls.Add(this.cardBudget, 0, 0);
            this.pageCapacity.Controls.Add(this.fieldCalculatorGroupBox, 0, 1);
            this.pageCapacity.Controls.Add(this.cardColumnCost, 0, 2);
            this.pageCapacity.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pageCapacity.Name = "pageCapacity";
            this.pageCapacity.RowCount = 3;
            this.pageCapacity.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 230F));
            this.pageCapacity.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 160F));
            this.pageCapacity.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.pageCapacity.Visible = false;
            //
            // cardBudget
            //
            this.cardBudget.Controls.Add(this.budgetLayout);
            this.cardBudget.Dock = System.Windows.Forms.DockStyle.Fill;
            this.cardBudget.Name = "cardBudget";
            this.cardBudget.Subtitle = "physical columns consumed in the database table";
            this.cardBudget.TabStop = false;
            this.cardBudget.Text = "Column budget";
            this.budgetLayout.ColumnCount = 1;
            this.budgetLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.budgetLayout.Controls.Add(this.budgetHeadline, 0, 0);
            this.budgetLayout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.budgetLayout.Margin = new System.Windows.Forms.Padding(0);
            this.budgetLayout.Name = "budgetLayout";
            this.budgetLayout.RowCount = 2;
            this.budgetLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 40F));
            this.budgetLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.budgetHeadline.Dock = System.Windows.Forms.DockStyle.Fill;
            this.budgetHeadline.Margin = new System.Windows.Forms.Padding(0);
            this.budgetHeadline.Name = "budgetHeadline";
            this.budgetHeadline.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            //
            // fieldCalculatorGroupBox
            //
            this.fieldCalculatorGroupBox.Controls.Add(this.tableLayoutPanel7);
            this.fieldCalculatorGroupBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.fieldCalculatorGroupBox.Name = "fieldCalculatorGroupBox";
            this.fieldCalculatorGroupBox.Subtitle = "type the fields you plan to add, the budget above updates live";
            this.fieldCalculatorGroupBox.TabIndex = 26;
            this.fieldCalculatorGroupBox.TabStop = false;
            this.fieldCalculatorGroupBox.Text = "Plan new fields";
            //
            // tableLayoutPanel7
            //
            this.tableLayoutPanel7.ColumnCount = 1;
            this.tableLayoutPanel7.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel7.Controls.Add(this.label5, 0, 0);
            this.tableLayoutPanel7.Controls.Add(this.tableLayoutPanel6, 0, 1);
            this.tableLayoutPanel7.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel7.Margin = new System.Windows.Forms.Padding(0);
            this.tableLayoutPanel7.Name = "tableLayoutPanel7";
            this.tableLayoutPanel7.RowCount = 2;
            this.tableLayoutPanel7.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 26F));
            this.tableLayoutPanel7.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            //
            // label5
            //
            this.label5.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label5.Margin = new System.Windows.Forms.Padding(0);
            this.label5.Name = "label5";
            this.label5.Text = "Each lookup uses 3 physical columns, each option set, money, boolean or status field uses 2, any other type uses 1. An entity is limited to 1,024 columns.";
            //
            // tableLayoutPanel6
            //
            this.tableLayoutPanel6.ColumnCount = 5;
            this.tableLayoutPanel6.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 150F));
            this.tableLayoutPanel6.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 260F));
            this.tableLayoutPanel6.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 150F));
            this.tableLayoutPanel6.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel6.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.AutoSize));
            this.tableLayoutPanel6.Controls.Add(this.label2, 0, 0);
            this.tableLayoutPanel6.Controls.Add(this.label3, 1, 0);
            this.tableLayoutPanel6.Controls.Add(this.label4, 2, 0);
            this.tableLayoutPanel6.Controls.Add(this.textBox1, 0, 1);
            this.tableLayoutPanel6.Controls.Add(this.textBox2, 1, 1);
            this.tableLayoutPanel6.Controls.Add(this.textBox3, 2, 1);
            this.tableLayoutPanel6.Controls.Add(this.calculatorResult, 3, 1);
            this.tableLayoutPanel6.Controls.Add(this.MoreDetailsLink, 4, 1);
            this.tableLayoutPanel6.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel6.Margin = new System.Windows.Forms.Padding(0);
            this.tableLayoutPanel6.Name = "tableLayoutPanel6";
            this.tableLayoutPanel6.RowCount = 2;
            this.tableLayoutPanel6.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 24F));
            this.tableLayoutPanel6.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 40F));
            //
            // calculator labels
            //
            this.label2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label2.Margin = new System.Windows.Forms.Padding(0);
            this.label2.Name = "label2";
            this.label2.Text = "Lookups";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.label3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label3.Margin = new System.Windows.Forms.Padding(0);
            this.label3.Name = "label3";
            this.label3.Text = "Option set, Money, Boolean, Status";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.label4.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label4.Margin = new System.Windows.Forms.Padding(0);
            this.label4.Name = "label4";
            this.label4.Text = "Other types";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            //
            // calculator inputs
            //
            this.textBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.textBox1.Margin = new System.Windows.Forms.Padding(0, 0, 12, 0);
            this.textBox1.MaxLength = 3;
            this.textBox1.Name = "textBox1";
            this.textBox1.NumericOnly = true;
            this.textBox1.Size = new System.Drawing.Size(138, 32);
            this.textBox1.TabIndex = 7;
            this.textBox1.Text = "0";
            this.textBox1.TextChanged += new System.EventHandler(this.calculatorInput_TextChanged);
            this.textBox2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.textBox2.Margin = new System.Windows.Forms.Padding(0, 0, 12, 0);
            this.textBox2.MaxLength = 3;
            this.textBox2.Name = "textBox2";
            this.textBox2.NumericOnly = true;
            this.textBox2.Size = new System.Drawing.Size(248, 32);
            this.textBox2.TabIndex = 8;
            this.textBox2.Text = "0";
            this.textBox2.TextChanged += new System.EventHandler(this.calculatorInput_TextChanged);
            this.textBox3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.textBox3.Margin = new System.Windows.Forms.Padding(0, 0, 12, 0);
            this.textBox3.MaxLength = 3;
            this.textBox3.Name = "textBox3";
            this.textBox3.NumericOnly = true;
            this.textBox3.Size = new System.Drawing.Size(138, 32);
            this.textBox3.TabIndex = 9;
            this.textBox3.Text = "0";
            this.textBox3.TextChanged += new System.EventHandler(this.calculatorInput_TextChanged);
            //
            // calculatorResult
            //
            this.calculatorResult.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.calculatorResult.AutoEllipsis = true;
            this.calculatorResult.Margin = new System.Windows.Forms.Padding(4, 0, 0, 0);
            this.calculatorResult.Name = "calculatorResult";
            this.calculatorResult.Size = new System.Drawing.Size(300, 32);
            this.calculatorResult.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            //
            // MoreDetailsLink
            //
            this.MoreDetailsLink.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.MoreDetailsLink.AutoSize = true;
            this.MoreDetailsLink.Cursor = System.Windows.Forms.Cursors.Hand;
            this.MoreDetailsLink.Margin = new System.Windows.Forms.Padding(8, 0, 0, 0);
            this.MoreDetailsLink.Name = "MoreDetailsLink";
            this.MoreDetailsLink.Text = "Learn more about field limits";
            this.MoreDetailsLink.Click += new System.EventHandler(this.label6_Click);
            //
            // cardColumnCost
            //
            this.cardColumnCost.Dock = System.Windows.Forms.DockStyle.Fill;
            this.cardColumnCost.Name = "cardColumnCost";
            this.cardColumnCost.Subtitle = "columns consumed per field type (fields × weight)";
            this.cardColumnCost.TabStop = false;
            this.cardColumnCost.Text = "What consumes the budget";
            //
            // MyPluginControl
            //
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.tableLayoutPanel1);
            this.Controls.Add(this.toolStripMenu);
            this.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Name = "MyPluginControl";
            this.Size = new System.Drawing.Size(1913, 1034);
            this.Load += new System.EventHandler(this.MyPluginControl_Load);
            this.toolStripMenu.ResumeLayout(false);
            this.toolStripMenu.PerformLayout();
            this.tableLayoutPanel1.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
            this.entitiesLayout.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.EntityGridView)).EndInit();
            this.rightLayout.ResumeLayout(false);
            this.pagesHost.ResumeLayout(false);
            this.pageWelcome.ResumeLayout(false);
            this.stepsLayout.ResumeLayout(false);
            this.guideActionPanel.ResumeLayout(false);
            this.guideActionPanel.PerformLayout();
            this.featuresLayout.ResumeLayout(false);
            this.pageOverview.ResumeLayout(false);
            this.kpiLayout.ResumeLayout(false);
            this.pageFields.ResumeLayout(false);
            this.fieldsGroupBox.ResumeLayout(false);
            this.tableLayoutPanel4.ResumeLayout(false);
            this.tableLayoutPanel5.ResumeLayout(false);
            this.tableLayoutPanel5.PerformLayout();
            this.fieldsFilterRow.ResumeLayout(false);
            this.fieldsFilterRow.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.fieldPropretiesView)).EndInit();
            this.pageUsage.ResumeLayout(false);
            this.usageLayout.ResumeLayout(false);
            this.usageKpiLayout.ResumeLayout(false);
            this.usageEmpty.ResumeLayout(false);
            this.usageEmptyLayout.ResumeLayout(false);
            this.pageCapacity.ResumeLayout(false);
            this.cardBudget.ResumeLayout(false);
            this.budgetLayout.ResumeLayout(false);
            this.fieldCalculatorGroupBox.ResumeLayout(false);
            this.tableLayoutPanel7.ResumeLayout(false);
            this.tableLayoutPanel6.ResumeLayout(false);
            this.tableLayoutPanel6.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ToolStrip toolStripMenu;
        private System.Windows.Forms.ToolStripSeparator tssSeparator1;
        private System.Windows.Forms.ToolStripButton tsbClose;
        private System.Windows.Forms.ToolStripButton byButton;
        private System.Windows.Forms.ToolStripButton helpButton;
        private ToolStripFluentComboBox entityTypeComboBox;
        private System.Windows.Forms.ToolStripButton toolStripButton;
        private System.Windows.Forms.ToolStripButton analyseButton;
        private System.Windows.Forms.ToolStripButton buttonExport;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private ToolStripFluentComboBox AnalyseType;
        private System.Windows.Forms.ToolStripButton buymeacoffeeicon;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private CardBox groupBox1;
        private System.Windows.Forms.TableLayoutPanel entitiesLayout;
        private FluentTextBox searchEntity;
        private System.Windows.Forms.DataGridView EntityGridView;
        private System.Windows.Forms.TableLayoutPanel rightLayout;
        private ContextHeader contextHeader;
        private PivotTabs pivotTabs;
        private System.Windows.Forms.Panel pagesHost;
        // Welcome
        private System.Windows.Forms.TableLayoutPanel pageWelcome;
        private System.Windows.Forms.Label welcomeTitle;
        private System.Windows.Forms.TableLayoutPanel stepsLayout;
        private StepCard stepLoad;
        private StepCard stepPick;
        private StepCard stepAnalyse;
        private System.Windows.Forms.FlowLayoutPanel guideActionPanel;
        private FluentButton guideAction;
        private System.Windows.Forms.Label guideHint;
        private System.Windows.Forms.Label featuresTitle;
        private System.Windows.Forms.TableLayoutPanel featuresLayout;
        private CardBox featureOverview;
        private System.Windows.Forms.Label featureOverviewText;
        private CardBox featureFields;
        private System.Windows.Forms.Label featureFieldsText;
        private CardBox featureUsage;
        private System.Windows.Forms.Label featureUsageText;
        private CardBox featureCapacity;
        private System.Windows.Forms.Label featureCapacityText;
        // Overview
        private System.Windows.Forms.TableLayoutPanel pageOverview;
        private System.Windows.Forms.TableLayoutPanel kpiLayout;
        private KpiTile kpiFields;
        private KpiTile kpiCustom;
        private KpiTile kpiUnmanaged;
        private KpiTile kpiCapacity;
        private CardBox cardTypes;
        private CardBox cardOrigin;
        private CardBox cardSettings;
        private CardBox cardPublishers;
        private CardBox cardTimeline;
        // Fields
        private System.Windows.Forms.Panel pageFields;
        private CardBox fieldsGroupBox;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel4;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel5;
        private FluentComboBox fieldTypeCombobox;
        private FluentTextBox searchField;
        private ToggleSwitch displayAllColumns;
        private System.Windows.Forms.TableLayoutPanel fieldsFilterRow;
        private ChipBar fieldsChips;
        private System.Windows.Forms.Label fieldsResultLabel;
        private System.Windows.Forms.DataGridView fieldPropretiesView;
        // Data usage
        private System.Windows.Forms.Panel pageUsage;
        private System.Windows.Forms.TableLayoutPanel usageLayout;
        private System.Windows.Forms.TableLayoutPanel usageKpiLayout;
        private KpiTile kpiRecords;
        private KpiTile kpiUnused;
        private KpiTile kpiLowUsage;
        private KpiTile kpiAverage;
        private CardBox cardFillDistribution;
        private CardBox cardFillByType;
        private CardBox cardLeastUsed;
        private CardBox usageEmpty;
        private System.Windows.Forms.TableLayoutPanel usageEmptyLayout;
        private System.Windows.Forms.Label usageEmptyText;
        private FluentButton rerunWithUsage;
        // Capacity
        private System.Windows.Forms.TableLayoutPanel pageCapacity;
        private CardBox cardBudget;
        private System.Windows.Forms.TableLayoutPanel budgetLayout;
        private System.Windows.Forms.Label budgetHeadline;
        private CardBox fieldCalculatorGroupBox;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel7;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel6;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private FluentTextBox textBox1;
        private FluentTextBox textBox2;
        private FluentTextBox textBox3;
        private System.Windows.Forms.Label calculatorResult;
        private System.Windows.Forms.Label MoreDetailsLink;
        private CardBox cardColumnCost;
    }
}
