using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace ExcelToYamlAddin
{
    partial class Ribbon
    {
        /// <summary>
        /// 디자이너 지원에 필요한 변수입니다.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// 사용 중인 모든 리소스를 정리합니다.
        /// </summary>
        /// <param name="disposing">관리되는 리소스를 삭제해야 하면 true이고, 그렇지 않으면 false입니다.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 구성 요소 디자이너에서 생성한 코드

        /// <summary>
        /// 디자이너 지원에 필요한 메서드입니다.
        /// 이 메서드의 내용을 코드 편집기로 수정하지 마세요.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Ribbon));
            this.tabExcelToJson = this.Factory.CreateRibbonTab();
            this.groupConvert = this.Factory.CreateRibbonGroup();
            this.groupSettings = this.Factory.CreateRibbonGroup();
            this.groupHelp = this.Factory.CreateRibbonGroup();
            this.btnConvertToYaml = this.Factory.CreateRibbonButton();
            this.btnSheetPathSettings = this.Factory.CreateRibbonButton();
            this.btnHelp = this.Factory.CreateRibbonButton();
            this.tabExcelToJson.SuspendLayout();
            this.groupConvert.SuspendLayout();
            this.groupSettings.SuspendLayout();
            this.groupHelp.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabExcelToJson
            // 
            this.tabExcelToJson.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabExcelToJson.Groups.Add(this.groupConvert);
            this.tabExcelToJson.Groups.Add(this.groupSettings);
            this.tabExcelToJson.Groups.Add(this.groupHelp);
            this.tabExcelToJson.Label = "Excel2Yaml";
            this.tabExcelToJson.Name = "tabExcelToJson";
            // 
            // groupConvert
            // 
            this.groupConvert.Items.Add(this.btnConvertToYaml);
            this.groupConvert.Label = "변환";
            this.groupConvert.Name = "groupConvert";
            // 
            // groupSettings
            // 
            this.groupSettings.Items.Add(this.btnSheetPathSettings);
            this.groupSettings.Label = "설정";
            this.groupSettings.Name = "groupSettings";
            // 
            // groupHelp
            // 
            this.groupHelp.Items.Add(this.btnHelp);
            this.groupHelp.Label = "도움말";
            this.groupHelp.Name = "groupHelp";
            // 
            // btnConvertToYaml
            // 
            this.btnConvertToYaml.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnConvertToYaml.Image = ((System.Drawing.Image)(resources.GetObject("btnConvertToYaml.Image")));
            this.btnConvertToYaml.Label = "YAML 변환";
            this.btnConvertToYaml.Name = "btnConvertToYaml";
            this.btnConvertToYaml.ScreenTip = "Excel을 YAML로 변환";
            this.btnConvertToYaml.ShowImage = true;
            this.btnConvertToYaml.SuperTip = "현재 워크시트의 데이터를 YAML 형식으로 변환합니다.";
            this.btnConvertToYaml.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnConvertToYamlClick);
            // 
            // btnSheetPathSettings
            // 
            this.btnSheetPathSettings.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnSheetPathSettings.Image = ((System.Drawing.Image)(resources.GetObject("btnSheetPathSettings.Image")));
            this.btnSheetPathSettings.Label = "설정";
            this.btnSheetPathSettings.Name = "btnSheetPathSettings";
            this.btnSheetPathSettings.ScreenTip = "시트별 경로 설정";
            this.btnSheetPathSettings.ShowImage = true;
            this.btnSheetPathSettings.SuperTip = "시트별로 저장 경로를 설정합니다. 각 시트마다 다른 경로에 저장할 수 있습니다.";
            this.btnSheetPathSettings.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnSheetPathSettingsClick);
            // 
            // btnHelp
            // 
            this.btnHelp.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnHelp.Image = ((System.Drawing.Image)(resources.GetObject("btnHelp.Image")));
            this.btnHelp.Label = "사용안내";
            this.btnHelp.Name = "btnHelp";
            this.btnHelp.ScreenTip = "Excel2YAML 사용 설명서";
            this.btnHelp.ShowImage = true;
            this.btnHelp.SuperTip = "Excel2YAML 사용 설명서를 엽니다.";
            this.btnHelp.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnHelpButtonClick);
            // 
            // Ribbon
            // 
            this.Name = "Ribbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tabExcelToJson);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Load);
            this.tabExcelToJson.ResumeLayout(false);
            this.tabExcelToJson.PerformLayout();
            this.groupConvert.ResumeLayout(false);
            this.groupConvert.PerformLayout();
            this.groupSettings.ResumeLayout(false);
            this.groupSettings.PerformLayout();
            this.groupHelp.ResumeLayout(false);
            this.groupHelp.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabExcelToJson;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupConvert;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnConvertToYaml;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupSettings;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSheetPathSettings;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupHelp;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnHelp;
    }
} 