using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace ExcelToJsonAddin
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
            this.tabExcelToJson = this.Factory.CreateRibbonTab();
            this.groupConvert = this.Factory.CreateRibbonGroup();
            this.btnConvertToJson = this.Factory.CreateRibbonButton();
            this.btnConvertToYaml = this.Factory.CreateRibbonButton();
            this.groupSettings = this.Factory.CreateRibbonGroup();
            this.btnSheetPathSettings = this.Factory.CreateRibbonButton();
            this.tabExcelToJson.SuspendLayout();
            this.groupConvert.SuspendLayout();
            this.groupSettings.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabExcelToJson
            // 
            this.tabExcelToJson.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabExcelToJson.Groups.Add(this.groupConvert);
            this.tabExcelToJson.Groups.Add(this.groupSettings);
            this.tabExcelToJson.Label = "Excel2Json";
            this.tabExcelToJson.Name = "tabExcelToJson";
            // 
            // groupConvert
            // 
            this.groupConvert.Items.Add(this.btnConvertToJson);
            this.groupConvert.Items.Add(this.btnConvertToYaml);
            this.groupConvert.Label = "변환";
            this.groupConvert.Name = "groupConvert";
            // 
            // btnConvertToJson
            // 
            this.btnConvertToJson.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnConvertToJson.Label = "JSON 변환";
            this.btnConvertToJson.Name = "btnConvertToJson";
            this.btnConvertToJson.ScreenTip = "Excel을 JSON으로 변환";
            this.btnConvertToJson.ShowImage = true;
            this.btnConvertToJson.SuperTip = "현재 워크시트의 데이터를 JSON 형식으로 변환합니다.";
            this.btnConvertToJson.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnConvertToJsonClick);
            // 
            // btnConvertToYaml
            // 
            this.btnConvertToYaml.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnConvertToYaml.Label = "YAML 변환";
            this.btnConvertToYaml.Name = "btnConvertToYaml";
            this.btnConvertToYaml.ScreenTip = "Excel을 YAML로 변환";
            this.btnConvertToYaml.ShowImage = true;
            this.btnConvertToYaml.SuperTip = "현재 워크시트의 데이터를 YAML 형식으로 변환합니다.";
            this.btnConvertToYaml.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnConvertToYamlClick);
            // 
            // groupSettings
            // 
            this.groupSettings.Items.Add(this.btnSheetPathSettings);
            this.groupSettings.Label = "설정";
            this.groupSettings.Name = "groupSettings";
            // 
            // btnSheetPathSettings
            // 
            this.btnSheetPathSettings.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnSheetPathSettings.Label = "시트별 경로 설정";
            this.btnSheetPathSettings.Name = "btnSheetPathSettings";
            this.btnSheetPathSettings.ScreenTip = "시트별 경로 설정";
            this.btnSheetPathSettings.ShowImage = true;
            this.btnSheetPathSettings.SuperTip = "시트별로 저장 경로를 설정합니다. 각 시트마다 다른 경로에 저장할 수 있습니다.";
            this.btnSheetPathSettings.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnSheetPathSettingsClick);
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
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabExcelToJson;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupConvert;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnConvertToJson;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnConvertToYaml;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupSettings;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSheetPathSettings;
    }
} 