using System.Collections.Generic;

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
        /// <param name="disposing">true si les ressources managées doivent être supprimées ; sinon, false.</param>
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
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.Windows.Forms.DataVisualization.Charting.ChartArea chartArea1 = new System.Windows.Forms.DataVisualization.Charting.ChartArea();
            System.Windows.Forms.DataVisualization.Charting.Legend legend1 = new System.Windows.Forms.DataVisualization.Charting.Legend();
            System.Windows.Forms.DataVisualization.Charting.Series series1 = new System.Windows.Forms.DataVisualization.Charting.Series();
            System.Windows.Forms.DataVisualization.Charting.ChartArea chartArea2 = new System.Windows.Forms.DataVisualization.Charting.ChartArea();
            System.Windows.Forms.DataVisualization.Charting.Legend legend2 = new System.Windows.Forms.DataVisualization.Charting.Legend();
            System.Windows.Forms.DataVisualization.Charting.Series series2 = new System.Windows.Forms.DataVisualization.Charting.Series();
            System.Windows.Forms.DataVisualization.Charting.ChartArea chartArea3 = new System.Windows.Forms.DataVisualization.Charting.ChartArea();
            System.Windows.Forms.DataVisualization.Charting.Legend legend3 = new System.Windows.Forms.DataVisualization.Charting.Legend();
            System.Windows.Forms.DataVisualization.Charting.Series series3 = new System.Windows.Forms.DataVisualization.Charting.Series();
            this.toolStripMenu = new System.Windows.Forms.ToolStrip();
            this.tssSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.tsbClose = new System.Windows.Forms.ToolStripButton();
            this.getEntitiesButton = new System.Windows.Forms.Button();
            this.EntityGridView = new System.Windows.Forms.DataGridView();
            this.visualStudioToolStripExtender1 = new WeifenLuo.WinFormsUI.Docking.VisualStudioToolStripExtender(this.components);
            this.entityTypeComboBox = new System.Windows.Forms.ComboBox();
            this.searchEntity = new System.Windows.Forms.TextBox();
            this.analyseButton = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.fieldPropretiesView = new System.Windows.Forms.DataGridView();
            this.label1 = new System.Windows.Forms.Label();
            this.fieldTypeCombobox = new System.Windows.Forms.ComboBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.ManagedUnmanagedFieldsgroupBox = new System.Windows.Forms.GroupBox();
            this.chartManagedUnmanagedFields = new System.Windows.Forms.DataVisualization.Charting.Chart();
            this.EntityFieldsCreatedGroupBox = new System.Windows.Forms.GroupBox();
            this.ChartFieldAvailable = new System.Windows.Forms.DataVisualization.Charting.Chart();
            this.EntityFieldsTypesGroupBox = new System.Windows.Forms.GroupBox();
            this.ChartFieldTypes = new System.Windows.Forms.DataVisualization.Charting.Chart();
            this.DisplayPercentageCheckbox = new System.Windows.Forms.CheckBox();
            this.fieldCalculatorGroupBox = new System.Windows.Forms.GroupBox();
            this.label7 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.textBox3 = new System.Windows.Forms.TextBox();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.button2 = new System.Windows.Forms.Button();
            this.toolStripMenu.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.EntityGridView)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.fieldPropretiesView)).BeginInit();
            this.groupBox3.SuspendLayout();
            this.ManagedUnmanagedFieldsgroupBox.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.chartManagedUnmanagedFields)).BeginInit();
            this.EntityFieldsCreatedGroupBox.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.ChartFieldAvailable)).BeginInit();
            this.EntityFieldsTypesGroupBox.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.ChartFieldTypes)).BeginInit();
            this.fieldCalculatorGroupBox.SuspendLayout();
            this.SuspendLayout();
            // 
            // toolStripMenu
            // 
            this.toolStripMenu.ImageScalingSize = new System.Drawing.Size(24, 24);
            this.toolStripMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tssSeparator1,
            this.tsbClose});
            this.toolStripMenu.Location = new System.Drawing.Point(0, 0);
            this.toolStripMenu.Name = "toolStripMenu";
            this.toolStripMenu.Size = new System.Drawing.Size(1913, 25);
            this.toolStripMenu.TabIndex = 4;
            this.toolStripMenu.Text = "toolStrip1";
            // 
            // tssSeparator1
            // 
            this.tssSeparator1.Name = "tssSeparator1";
            this.tssSeparator1.Size = new System.Drawing.Size(6, 25);
            // 
            // tsbClose
            // 
            this.tsbClose.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.tsbClose.Name = "tsbClose";
            this.tsbClose.Size = new System.Drawing.Size(86, 22);
            this.tsbClose.Text = "Close this tool";
            this.tsbClose.Click += new System.EventHandler(this.tsbClose_Click);
            // 
            // getEntitiesButton
            // 
            this.getEntitiesButton.Location = new System.Drawing.Point(231, 2);
            this.getEntitiesButton.Name = "getEntitiesButton";
            this.getEntitiesButton.Size = new System.Drawing.Size(75, 23);
            this.getEntitiesButton.TabIndex = 6;
            this.getEntitiesButton.Text = "Get Entities";
            this.getEntitiesButton.UseVisualStyleBackColor = true;
            this.getEntitiesButton.Click += new System.EventHandler(this.getEntitiesButton_Click);
            // 
            // EntityGridView
            // 
            this.EntityGridView.AllowUserToAddRows = false;
            this.EntityGridView.AllowUserToDeleteRows = false;
            this.EntityGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.EntityGridView.Dock = System.Windows.Forms.DockStyle.Left;
            this.EntityGridView.Location = new System.Drawing.Point(3, 16);
            this.EntityGridView.MultiSelect = false;
            this.EntityGridView.Name = "EntityGridView";
            this.EntityGridView.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.EntityGridView.Size = new System.Drawing.Size(482, 990);
            this.EntityGridView.TabIndex = 3;
            this.EntityGridView.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.EntityGridView_CellContentClick_1);
            // 
            // visualStudioToolStripExtender1
            // 
            this.visualStudioToolStripExtender1.DefaultRenderer = null;
            // 
            // entityTypeComboBox
            // 
            this.entityTypeComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.entityTypeComboBox.Items.AddRange(new object[] {
            "All Entities",
            "Custom Entities",
            "CRM Entities",
            "System Entities"});
            this.entityTypeComboBox.Location = new System.Drawing.Point(104, 3);
            this.entityTypeComboBox.Name = "entityTypeComboBox";
            this.entityTypeComboBox.Size = new System.Drawing.Size(121, 21);
            this.entityTypeComboBox.TabIndex = 3;
            // 
            // searchEntity
            // 
            this.searchEntity.Location = new System.Drawing.Point(310, 3);
            this.searchEntity.Name = "searchEntity";
            this.searchEntity.Size = new System.Drawing.Size(100, 20);
            this.searchEntity.TabIndex = 1;
            this.searchEntity.Text = "Search";
            this.searchEntity.Click += new System.EventHandler(this.textBox1_Click);
            this.searchEntity.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
            // 
            // analyseButton
            // 
            this.analyseButton.Location = new System.Drawing.Point(416, 2);
            this.analyseButton.Name = "analyseButton";
            this.analyseButton.Size = new System.Drawing.Size(75, 23);
            this.analyseButton.TabIndex = 9;
            this.analyseButton.Text = "Analyse";
            this.analyseButton.UseVisualStyleBackColor = true;
            this.analyseButton.Click += new System.EventHandler(this.analyseButton_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.EntityGridView);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Left;
            this.groupBox1.Location = new System.Drawing.Point(0, 25);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(498, 1009);
            this.groupBox1.TabIndex = 10;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Entities";
            // 
            // groupBox2
            // 
            this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox2.Controls.Add(this.fieldPropretiesView);
            this.groupBox2.Controls.Add(this.label1);
            this.groupBox2.Controls.Add(this.fieldTypeCombobox);
            this.groupBox2.Location = new System.Drawing.Point(504, 25);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(1409, 344);
            this.groupBox2.TabIndex = 11;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Entity Analysis";
            // 
            // fieldPropretiesView
            // 
            this.fieldPropretiesView.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.fieldPropretiesView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.fieldPropretiesView.Location = new System.Drawing.Point(6, 46);
            this.fieldPropretiesView.Name = "fieldPropretiesView";
            this.fieldPropretiesView.Size = new System.Drawing.Size(1385, 295);
            this.fieldPropretiesView.TabIndex = 2;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 22);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(61, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Fields Type";
            // 
            // fieldTypeCombobox
            // 
            this.fieldTypeCombobox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.fieldTypeCombobox.FormattingEnabled = true;
            this.fieldTypeCombobox.Location = new System.Drawing.Point(73, 19);
            this.fieldTypeCombobox.Name = "fieldTypeCombobox";
            this.fieldTypeCombobox.Size = new System.Drawing.Size(319, 21);
            this.fieldTypeCombobox.TabIndex = 0;
            this.fieldTypeCombobox.SelectedIndexChanged += new System.EventHandler(this.fieldTypeCombobox_SelectedIndexChanged);
            // 
            // groupBox3
            // 
            this.groupBox3.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox3.BackColor = System.Drawing.SystemColors.Window;
            this.groupBox3.Controls.Add(this.ManagedUnmanagedFieldsgroupBox);
            this.groupBox3.Controls.Add(this.EntityFieldsCreatedGroupBox);
            this.groupBox3.Controls.Add(this.EntityFieldsTypesGroupBox);
            this.groupBox3.Controls.Add(this.DisplayPercentageCheckbox);
            this.groupBox3.Location = new System.Drawing.Point(507, 372);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(1391, 527);
            this.groupBox3.TabIndex = 12;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Reports";
            // 
            // ManagedUnmanagedFieldsgroupBox
            // 
            this.ManagedUnmanagedFieldsgroupBox.Controls.Add(this.chartManagedUnmanagedFields);
            this.ManagedUnmanagedFieldsgroupBox.Location = new System.Drawing.Point(465, 41);
            this.ManagedUnmanagedFieldsgroupBox.Name = "ManagedUnmanagedFieldsgroupBox";
            this.ManagedUnmanagedFieldsgroupBox.Size = new System.Drawing.Size(450, 375);
            this.ManagedUnmanagedFieldsgroupBox.TabIndex = 6;
            this.ManagedUnmanagedFieldsgroupBox.TabStop = false;
            this.ManagedUnmanagedFieldsgroupBox.Text = "Managed/UnmanagedFields";
            this.ManagedUnmanagedFieldsgroupBox.Visible = false;
            // 
            // chartManagedUnmanagedFields
            // 
            chartArea1.Name = "ChartArea1";
            this.chartManagedUnmanagedFields.ChartAreas.Add(chartArea1);
            this.chartManagedUnmanagedFields.Dock = System.Windows.Forms.DockStyle.Fill;
            legend1.Name = "Legend1";
            this.chartManagedUnmanagedFields.Legends.Add(legend1);
            this.chartManagedUnmanagedFields.Location = new System.Drawing.Point(3, 16);
            this.chartManagedUnmanagedFields.Name = "chartManagedUnmanagedFields";
            series1.ChartArea = "ChartArea1";
            series1.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Doughnut;
            series1.Legend = "Legend1";
            series1.Name = "managedUnmanagedFields";
            this.chartManagedUnmanagedFields.Series.Add(series1);
            this.chartManagedUnmanagedFields.Size = new System.Drawing.Size(444, 356);
            this.chartManagedUnmanagedFields.TabIndex = 3;
            // 
            // EntityFieldsCreatedGroupBox
            // 
            this.EntityFieldsCreatedGroupBox.Controls.Add(this.ChartFieldAvailable);
            this.EntityFieldsCreatedGroupBox.Location = new System.Drawing.Point(5, 41);
            this.EntityFieldsCreatedGroupBox.Name = "EntityFieldsCreatedGroupBox";
            this.EntityFieldsCreatedGroupBox.Size = new System.Drawing.Size(450, 375);
            this.EntityFieldsCreatedGroupBox.TabIndex = 5;
            this.EntityFieldsCreatedGroupBox.TabStop = false;
            this.EntityFieldsCreatedGroupBox.Text = "Entity Fields Created";
            this.EntityFieldsCreatedGroupBox.Visible = false;
            // 
            // ChartFieldAvailable
            // 
            chartArea2.Name = "ChartArea1";
            this.ChartFieldAvailable.ChartAreas.Add(chartArea2);
            this.ChartFieldAvailable.Dock = System.Windows.Forms.DockStyle.Fill;
            legend2.Name = "Legend1";
            this.ChartFieldAvailable.Legends.Add(legend2);
            this.ChartFieldAvailable.Location = new System.Drawing.Point(3, 16);
            this.ChartFieldAvailable.Name = "ChartFieldAvailable";
            series2.ChartArea = "ChartArea1";
            series2.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Doughnut;
            series2.Legend = "Legend1";
            series2.Name = "AvailableField";
            this.ChartFieldAvailable.Series.Add(series2);
            this.ChartFieldAvailable.Size = new System.Drawing.Size(444, 356);
            this.ChartFieldAvailable.TabIndex = 1;
            // 
            // EntityFieldsTypesGroupBox
            // 
            this.EntityFieldsTypesGroupBox.Controls.Add(this.ChartFieldTypes);
            this.EntityFieldsTypesGroupBox.Location = new System.Drawing.Point(925, 41);
            this.EntityFieldsTypesGroupBox.Name = "EntityFieldsTypesGroupBox";
            this.EntityFieldsTypesGroupBox.Size = new System.Drawing.Size(450, 375);
            this.EntityFieldsTypesGroupBox.TabIndex = 4;
            this.EntityFieldsTypesGroupBox.TabStop = false;
            this.EntityFieldsTypesGroupBox.Text = "Entity Fields Types";
            this.EntityFieldsTypesGroupBox.Visible = false;
            // 
            // ChartFieldTypes
            // 
            chartArea3.Name = "ChartArea1";
            this.ChartFieldTypes.ChartAreas.Add(chartArea3);
            this.ChartFieldTypes.Dock = System.Windows.Forms.DockStyle.Fill;
            legend3.ItemColumnSpacing = 0;
            legend3.Name = "Legend1";
            this.ChartFieldTypes.Legends.Add(legend3);
            this.ChartFieldTypes.Location = new System.Drawing.Point(3, 16);
            this.ChartFieldTypes.Name = "ChartFieldTypes";
            series3.ChartArea = "ChartArea1";
            series3.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Doughnut;
            series3.Legend = "Legend1";
            series3.Name = "fieldsReport";
            series3.SmartLabelStyle.MaxMovingDistance = 0D;
            this.ChartFieldTypes.Series.Add(series3);
            this.ChartFieldTypes.Size = new System.Drawing.Size(444, 356);
            this.ChartFieldTypes.TabIndex = 0;
            // 
            // DisplayPercentageCheckbox
            // 
            this.DisplayPercentageCheckbox.AutoSize = true;
            this.DisplayPercentageCheckbox.Location = new System.Drawing.Point(7, 19);
            this.DisplayPercentageCheckbox.Name = "DisplayPercentageCheckbox";
            this.DisplayPercentageCheckbox.Size = new System.Drawing.Size(118, 17);
            this.DisplayPercentageCheckbox.TabIndex = 2;
            this.DisplayPercentageCheckbox.Text = "Display Percentage";
            this.DisplayPercentageCheckbox.UseVisualStyleBackColor = true;
            this.DisplayPercentageCheckbox.Visible = false;
            this.DisplayPercentageCheckbox.CheckedChanged += new System.EventHandler(this.DisplayPercentageCheckbox_CheckedChanged);
            // 
            // fieldCalculatorGroupBox
            // 
            this.fieldCalculatorGroupBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.fieldCalculatorGroupBox.BackColor = System.Drawing.SystemColors.Window;
            this.fieldCalculatorGroupBox.Controls.Add(this.label7);
            this.fieldCalculatorGroupBox.Controls.Add(this.button1);
            this.fieldCalculatorGroupBox.Controls.Add(this.label5);
            this.fieldCalculatorGroupBox.Controls.Add(this.label4);
            this.fieldCalculatorGroupBox.Controls.Add(this.label3);
            this.fieldCalculatorGroupBox.Controls.Add(this.label2);
            this.fieldCalculatorGroupBox.Controls.Add(this.textBox3);
            this.fieldCalculatorGroupBox.Controls.Add(this.textBox2);
            this.fieldCalculatorGroupBox.Controls.Add(this.textBox1);
            this.fieldCalculatorGroupBox.Location = new System.Drawing.Point(510, 917);
            this.fieldCalculatorGroupBox.Name = "fieldCalculatorGroupBox";
            this.fieldCalculatorGroupBox.Size = new System.Drawing.Size(1385, 106);
            this.fieldCalculatorGroupBox.TabIndex = 13;
            this.fieldCalculatorGroupBox.TabStop = false;
            this.fieldCalculatorGroupBox.Text = "Field I Can Create";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(1208, 84);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(0, 13);
            this.label7.TabIndex = 14;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(374, 55);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 7;
            this.button1.Text = "Verify";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(17, 16);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(245, 13);
            this.label5.TabIndex = 6;
            this.label5.Text = "Enter The Number Of Fields You Want To Create :";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(17, 84);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(65, 13);
            this.label4.TabIndex = 5;
            this.label4.Text = "Other Types";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(17, 58);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(163, 13);
            this.label3.TabIndex = 4;
            this.label3.Text = "PickList, Money, Boolean, Status";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(17, 35);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(80, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "Lookup, Owner";
            // 
            // textBox3
            // 
            this.textBox3.Location = new System.Drawing.Point(199, 81);
            this.textBox3.MaxLength = 3;
            this.textBox3.Name = "textBox3";
            this.textBox3.Size = new System.Drawing.Size(107, 20);
            this.textBox3.TabIndex = 2;
            this.textBox3.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.textBox3.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textBox3_KeyPress);
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(199, 57);
            this.textBox2.MaxLength = 3;
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(107, 20);
            this.textBox2.TabIndex = 1;
            this.textBox2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.textBox2.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textBox2_KeyPress);
            // 
            // textBox1
            // 
            this.textBox1.BackColor = System.Drawing.SystemColors.Window;
            this.textBox1.Location = new System.Drawing.Point(199, 32);
            this.textBox1.MaxLength = 3;
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(107, 20);
            this.textBox1.TabIndex = 0;
            this.textBox1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.textBox1.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textBox1_KeyPress);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(513, 2);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 23);
            this.button2.TabIndex = 14;
            this.button2.Text = "Export";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // MyPluginControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.button2);
            this.Controls.Add(this.fieldCalculatorGroupBox);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.searchEntity);
            this.Controls.Add(this.analyseButton);
            this.Controls.Add(this.entityTypeComboBox);
            this.Controls.Add(this.getEntitiesButton);
            this.Controls.Add(this.toolStripMenu);
            this.Name = "MyPluginControl";
            this.Size = new System.Drawing.Size(1913, 1034);
            this.Load += new System.EventHandler(this.MyPluginControl_Load);
            this.toolStripMenu.ResumeLayout(false);
            this.toolStripMenu.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.EntityGridView)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.fieldPropretiesView)).EndInit();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.ManagedUnmanagedFieldsgroupBox.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.chartManagedUnmanagedFields)).EndInit();
            this.EntityFieldsCreatedGroupBox.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.ChartFieldAvailable)).EndInit();
            this.EntityFieldsTypesGroupBox.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.ChartFieldTypes)).EndInit();
            this.fieldCalculatorGroupBox.ResumeLayout(false);
            this.fieldCalculatorGroupBox.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.ToolStrip toolStripMenu;
        private System.Windows.Forms.ToolStripSeparator tssSeparator1;
        private System.Windows.Forms.ToolStripButton tsbClose;
        private System.Windows.Forms.Button getEntitiesButton;
        private System.Windows.Forms.DataGridView EntityGridView;
        private WeifenLuo.WinFormsUI.Docking.VisualStudioToolStripExtender visualStudioToolStripExtender1;
        private System.Windows.Forms.ComboBox entityTypeComboBox;
        private System.Windows.Forms.TextBox searchEntity;
        private System.Windows.Forms.Button analyseButton;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox fieldTypeCombobox;
        private System.Windows.Forms.DataGridView fieldPropretiesView;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.GroupBox fieldCalculatorGroupBox;
        private System.Windows.Forms.DataVisualization.Charting.Chart ChartFieldAvailable;
        private System.Windows.Forms.DataVisualization.Charting.Chart ChartFieldTypes;
        private System.Windows.Forms.CheckBox DisplayPercentageCheckbox;
        private System.Windows.Forms.DataVisualization.Charting.Chart chartManagedUnmanagedFields;
        private System.Windows.Forms.GroupBox ManagedUnmanagedFieldsgroupBox;
        private System.Windows.Forms.GroupBox EntityFieldsCreatedGroupBox;
        private System.Windows.Forms.GroupBox EntityFieldsTypesGroupBox;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textBox3;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Button button2;
    }
}
