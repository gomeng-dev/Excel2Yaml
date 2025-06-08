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
            this.tabExcelToYaml = this.Factory.CreateRibbonTab();
            this.groupConvert = this.Factory.CreateRibbonGroup();
            this.btnConvertToYaml = this.Factory.CreateRibbonButton();
            this.btnConvertYamlToJson = this.Factory.CreateRibbonButton();
            this.btnConvertToXml = this.Factory.CreateRibbonButton(); // XML 버튼 추가
            this.groupTools = this.Factory.CreateRibbonGroup(); // 도구 그룹 추가
            this.btnImportXml = this.Factory.CreateRibbonButton(); // XML 가져오기 버튼 추가
            this.btnImportYaml = this.Factory.CreateRibbonButton(); // YAML 가져오기 버튼 추가
            this.btnExportToHtml = this.Factory.CreateRibbonButton(); // HTML 내보내기 버튼 추가
            this.btnTestXmlToYaml = this.Factory.CreateRibbonButton(); // XML to YAML 테스트 버튼 추가
            this.groupSettings = this.Factory.CreateRibbonGroup();
            this.btnSheetPathSettings = this.Factory.CreateRibbonButton();
            this.groupHelp = this.Factory.CreateRibbonGroup();
            this.btnHelp = this.Factory.CreateRibbonButton();
            this.tabExcelToYaml.SuspendLayout();
            this.groupConvert.SuspendLayout();
            this.groupTools.SuspendLayout();
            this.groupSettings.SuspendLayout();
            this.groupHelp.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabExcelToYaml
            // 
            this.tabExcelToYaml.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabExcelToYaml.Groups.Add(this.groupConvert);
            this.tabExcelToYaml.Groups.Add(this.groupTools);
            this.tabExcelToYaml.Groups.Add(this.groupSettings);
            this.tabExcelToYaml.Groups.Add(this.groupHelp);
            this.tabExcelToYaml.Label = "Excel2Yaml";
            this.tabExcelToYaml.Name = "tabExcelToYaml";
            // 
            // groupConvert
            // 
            this.groupConvert.Items.Add(this.btnConvertToYaml);
            this.groupConvert.Items.Add(this.btnConvertYamlToJson);
            this.groupConvert.Items.Add(this.btnConvertToXml); // XML 버튼 그룹에 추가
            this.groupConvert.Label = "변환";
            this.groupConvert.Name = "groupConvert";
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
            // btnConvertYamlToJson
            // 
            this.btnConvertYamlToJson.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnConvertYamlToJson.Image = ((System.Drawing.Image)(resources.GetObject("btnConvertYamlToJson.Image")));
            this.btnConvertYamlToJson.Label = "JSON 변환";
            this.btnConvertYamlToJson.Name = "btnConvertYamlToJson";
            this.btnConvertYamlToJson.ScreenTip = "Excel을 JSON으로 변환";
            this.btnConvertYamlToJson.ShowImage = true;
            this.btnConvertYamlToJson.SuperTip = "현재 워크시트의 데이터를 JSON 형식으로 변환합니다.";
            this.btnConvertYamlToJson.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnConvertYamlToJsonClick);
            // 
            // btnConvertToXml
            // 
            this.btnConvertToXml.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnConvertToXml.Image = ((System.Drawing.Image)(resources.GetObject("btnConvertToXml.Image"))); // 적절한 아이콘으로 변경 필요
            this.btnConvertToXml.Label = "XML 변환";
            this.btnConvertToXml.Name = "btnConvertToXml";
            this.btnConvertToXml.ScreenTip = "Excel을 XML로 변환 (YAML 경유)";
            this.btnConvertToXml.ShowImage = true;
            this.btnConvertToXml.SuperTip = "현재 워크시트의 데이터를 YAML로 변환 후, 그 결과를 XML 형식으로 변환합니다.";
            this.btnConvertToXml.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnConvertToXmlClick);
            // 
            // btnImportXml
            // 
            this.btnImportXml.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnImportXml.Image = ((System.Drawing.Image)(resources.GetObject("btnHelp.Image"))); // 임시로 Help 아이콘 사용
            this.btnImportXml.Label = "XML 가져오기";
            this.btnImportXml.Name = "btnImportXml";
            this.btnImportXml.ScreenTip = "XML을 Excel로 가져오기";
            this.btnImportXml.ShowImage = true;
            this.btnImportXml.SuperTip = "XML 파일을 읽어서 Excel 시트로 변환합니다.";
            this.btnImportXml.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnImportXmlClick);
            // 
            // btnImportYaml
            // 
            this.btnImportYaml.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnImportYaml.Image = ((System.Drawing.Image)(resources.GetObject("btnHelp.Image"))); // 임시로 Help 아이콘 사용
            this.btnImportYaml.Label = "YAML 가져오기";
            this.btnImportYaml.Name = "btnImportYaml";
            this.btnImportYaml.ScreenTip = "YAML을 Excel로 가져오기";
            this.btnImportYaml.ShowImage = true;
            this.btnImportYaml.SuperTip = "YAML 파일을 읽어서 Excel 시트로 변환합니다.";
            this.btnImportYaml.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnImportYamlClick);
            // 
            // btnExportToHtml
            // 
            this.btnExportToHtml.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnExportToHtml.Image = ((System.Drawing.Image)(resources.GetObject("btnHelp.Image"))); // 임시로 Help 아이콘 사용
            this.btnExportToHtml.Label = "HTML 내보내기";
            this.btnExportToHtml.Name = "btnExportToHtml";
            this.btnExportToHtml.ScreenTip = "현재 시트를 HTML로 내보내기";
            this.btnExportToHtml.ShowImage = true;
            this.btnExportToHtml.SuperTip = "현재 시트를 HTML 파일로 내보내서 브라우저에서 확인할 수 있습니다.";
            this.btnExportToHtml.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnExportToHtmlClick);
            // 
            // btnTestXmlToYaml
            // 
            this.btnTestXmlToYaml.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnTestXmlToYaml.Image = ((System.Drawing.Image)(resources.GetObject("btnHelp.Image"))); // 임시로 Help 아이콘 사용
            this.btnTestXmlToYaml.Label = "XML→YAML 테스트";
            this.btnTestXmlToYaml.Name = "btnTestXmlToYaml";
            this.btnTestXmlToYaml.ScreenTip = "XML을 YAML로 변환 테스트";
            this.btnTestXmlToYaml.ShowImage = true;
            this.btnTestXmlToYaml.SuperTip = "샘플 XML을 YAML로 변환하고 콘솔에 결과를 출력합니다.";
            this.btnTestXmlToYaml.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnTestXmlToYamlClick);
            // 
            // groupTools
            // 
            this.groupTools.Items.Add(this.btnImportXml);
            this.groupTools.Items.Add(this.btnImportYaml);
            this.groupTools.Items.Add(this.btnExportToHtml);
            this.groupTools.Items.Add(this.btnTestXmlToYaml);
            this.groupTools.Label = "도구";
            this.groupTools.Name = "groupTools";
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
            this.btnSheetPathSettings.Image = ((System.Drawing.Image)(resources.GetObject("btnSheetPathSettings.Image")));
            this.btnSheetPathSettings.Label = "설정";
            this.btnSheetPathSettings.Name = "btnSheetPathSettings";
            this.btnSheetPathSettings.ScreenTip = "시트별 경로 설정";
            this.btnSheetPathSettings.ShowImage = true;
            this.btnSheetPathSettings.SuperTip = "시트별로 저장 경로를 설정합니다. 각 시트마다 다른 경로에 저장할 수 있습니다.";
            this.btnSheetPathSettings.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnSheetPathSettingsClick);
            // 
            // groupHelp
            // 
            this.groupHelp.Items.Add(this.btnHelp);
            this.groupHelp.Label = "도움말";
            this.groupHelp.Name = "groupHelp";
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
            this.Tabs.Add(this.tabExcelToYaml);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Load);
            this.tabExcelToYaml.ResumeLayout(false);
            this.tabExcelToYaml.PerformLayout();
            this.groupConvert.ResumeLayout(false);
            this.groupConvert.PerformLayout();
            this.groupTools.ResumeLayout(false);
            this.groupTools.PerformLayout();
            this.groupSettings.ResumeLayout(false);
            this.groupSettings.PerformLayout();
            this.groupHelp.ResumeLayout(false);
            this.groupHelp.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabExcelToYaml;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupConvert;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnConvertToYaml;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnConvertYamlToJson;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnConvertToXml; // XML 버튼 멤버 추가
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupTools; // 도구 그룹 멤버 추가
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnImportXml; // XML 가져오기 버튼 멤버 추가
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnImportYaml; // YAML 가져오기 버튼 멤버 추가
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnExportToHtml; // HTML 내보내기 버튼 멤버 추가
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnTestXmlToYaml; // XML to YAML 테스트 버튼 멤버 추가
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupSettings;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSheetPathSettings;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupHelp;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnHelp;
    }
} 