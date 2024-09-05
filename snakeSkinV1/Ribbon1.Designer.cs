namespace snakeSkinV1
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該處置 Managed 資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 元件設計工具產生的程式碼

        /// <summary>
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器修改
        /// 這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl2 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl3 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl4 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl5 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl6 = this.Factory.CreateRibbonDropDownItem();
            this.tab1 = this.Factory.CreateRibbonTab();
            this.dcOperate = this.Factory.CreateRibbonGroup();
            this.capture = this.Factory.CreateRibbonButton();
            this.addOne = this.Factory.CreateRibbonButton();
            this.addSTVmode = this.Factory.CreateRibbonMenu();
            this.sourceSelectMode = this.Factory.CreateRibbonCheckBox();
            this.targetSelectMode = this.Factory.CreateRibbonCheckBox();
            this.valueSelectMode = this.Factory.CreateRibbonCheckBox();
            this.removeSelection = this.Factory.CreateRibbonDropDown();
            this.doRemoveSelection = this.Factory.CreateRibbonButton();
            this.settingDC = this.Factory.CreateRibbonMenu();
            this.autoNextPT = this.Factory.CreateRibbonCheckBox();
            this.autoPreView = this.Factory.CreateRibbonCheckBox();
            this.hotfixAutoReset41 = this.Factory.CreateRibbonCheckBox();
            this.safe3 = this.Factory.CreateRibbonCheckBox();
            this.kpOperate = this.Factory.CreateRibbonGroup();
            this.arraySetSource = this.Factory.CreateRibbonButton();
            this.arraySetTarget = this.Factory.CreateRibbonButton();
            this.previewArray = this.Factory.CreateRibbonButton();
            this.arraySetData = this.Factory.CreateRibbonButton();
            this.arrayColorSetting = this.Factory.CreateRibbonMenu();
            this.displayColorAfterSelect = this.Factory.CreateRibbonCheckBox();
            this.picColor1 = this.Factory.CreateRibbonButton();
            this.picColor2 = this.Factory.CreateRibbonButton();
            this.picColor3 = this.Factory.CreateRibbonButton();
            this.picColor4 = this.Factory.CreateRibbonButton();
            this.picColor5 = this.Factory.CreateRibbonButton();
            this.picColor6 = this.Factory.CreateRibbonButton();
            this.primaryOP = this.Factory.CreateRibbonGroup();
            this.juniorOP = this.Factory.CreateRibbonGroup();
            this.addSplitButton = this.Factory.CreateRibbonButton();
            this.testAddRow = this.Factory.CreateRibbonButton();
            this.seniorOP = this.Factory.CreateRibbonGroup();
            this.mgOperate = this.Factory.CreateRibbonGroup();
            this.modeEdit = this.Factory.CreateRibbonDropDown();
            this.editData = this.Factory.CreateRibbonGallery();
            this.clearVisual = this.Factory.CreateRibbonButton();
            this.mgOperateColorSetting = this.Factory.CreateRibbonMenu();
            this.a1show = this.Factory.CreateRibbonButton();
            this.a2show = this.Factory.CreateRibbonButton();
            this.b1show = this.Factory.CreateRibbonButton();
            this.b2show = this.Factory.CreateRibbonButton();
            this.c1show = this.Factory.CreateRibbonButton();
            this.c2show = this.Factory.CreateRibbonButton();
            this.rainbowMG = this.Factory.CreateRibbonButton();
            this.dpOperate = this.Factory.CreateRibbonGroup();
            this.displayData = this.Factory.CreateRibbonButton();
            this.processData = this.Factory.CreateRibbonButton();
            this.plotTitle = this.Factory.CreateRibbonEditBox();
            this.autoNodeColorSetting = this.Factory.CreateRibbonToggleButton();
            this.ucOperate = this.Factory.CreateRibbonGroup();
            this.saveMap = this.Factory.CreateRibbonButton();
            this.loadMap = this.Factory.CreateRibbonButton();
            this.exportMap = this.Factory.CreateRibbonButton();
            this.importMap = this.Factory.CreateRibbonButton();
            this.setting4save = this.Factory.CreateRibbonMenu();
            this.emptyWhenLoad = this.Factory.CreateRibbonCheckBox();
            this.ableShift = this.Factory.CreateRibbonCheckBox();
            this.shiftSetting = this.Factory.CreateRibbonGallery();
            this.pdOperate = this.Factory.CreateRibbonGroup();
            this.debugHide = this.Factory.CreateRibbonMenu();
            this.addMainData = this.Factory.CreateRibbonButton();
            this.readUserSelectB = this.Factory.CreateRibbonButton();
            this.addRibbonDropdownItemB = this.Factory.CreateRibbonButton();
            this.listTest = this.Factory.CreateRibbonButton();
            this.writeMainDataDumb = this.Factory.CreateRibbonButton();
            this.galleryNumTest = this.Factory.CreateRibbonButton();
            this.rainbowTest = this.Factory.CreateRibbonButton();
            this.testsave = this.Factory.CreateRibbonButton();
            this.testloadsave = this.Factory.CreateRibbonButton();
            this.worksheetcodenametest = this.Factory.CreateRibbonButton();
            this.testActivateWindows = this.Factory.CreateRibbonButton();
            this.newWindowsTag = this.Factory.CreateRibbonButton();
            this.defaultSnakeColorTest = this.Factory.CreateRibbonButton();
            this.todolist = this.Factory.CreateRibbonButton();
            this.Rpath = this.Factory.CreateRibbonEditBox();
            this.musicPath = this.Factory.CreateRibbonEditBox();
            this.useOldR = this.Factory.CreateRibbonToggleButton();
            this.a1 = new System.Windows.Forms.ColorDialog();
            this.a2 = new System.Windows.Forms.ColorDialog();
            this.b1 = new System.Windows.Forms.ColorDialog();
            this.b2 = new System.Windows.Forms.ColorDialog();
            this.c1 = new System.Windows.Forms.ColorDialog();
            this.c2 = new System.Windows.Forms.ColorDialog();
            this.arrayColorSetSource1 = new System.Windows.Forms.ColorDialog();
            this.arrayColorSetSource2 = new System.Windows.Forms.ColorDialog();
            this.arrayColorSetTarget1 = new System.Windows.Forms.ColorDialog();
            this.arrayColorSetTarget2 = new System.Windows.Forms.ColorDialog();
            this.arrayColorSetData1 = new System.Windows.Forms.ColorDialog();
            this.arrayColorSetData2 = new System.Windows.Forms.ColorDialog();
            this.saveMirrorText = new System.Windows.Forms.SaveFileDialog();
            this.loadMirrorText = new System.Windows.Forms.OpenFileDialog();
            this.assembHTML = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.dcOperate.SuspendLayout();
            this.kpOperate.SuspendLayout();
            this.juniorOP.SuspendLayout();
            this.mgOperate.SuspendLayout();
            this.dpOperate.SuspendLayout();
            this.ucOperate.SuspendLayout();
            this.pdOperate.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.dcOperate);
            this.tab1.Groups.Add(this.kpOperate);
            this.tab1.Groups.Add(this.primaryOP);
            this.tab1.Groups.Add(this.juniorOP);
            this.tab1.Groups.Add(this.seniorOP);
            this.tab1.Groups.Add(this.mgOperate);
            this.tab1.Groups.Add(this.dpOperate);
            this.tab1.Groups.Add(this.ucOperate);
            this.tab1.Groups.Add(this.pdOperate);
            this.tab1.Label = "蛇皮圖V1";
            this.tab1.Name = "tab1";
            // 
            // dcOperate
            // 
            this.dcOperate.Items.Add(this.capture);
            this.dcOperate.Items.Add(this.addOne);
            this.dcOperate.Items.Add(this.addSTVmode);
            this.dcOperate.Items.Add(this.removeSelection);
            this.dcOperate.Items.Add(this.doRemoveSelection);
            this.dcOperate.Items.Add(this.settingDC);
            this.dcOperate.Label = "dcOperate";
            this.dcOperate.Name = "dcOperate";
            // 
            // capture
            // 
            this.capture.Label = "擷取";
            this.capture.Name = "capture";
            this.capture.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.capture_Click);
            // 
            // addOne
            // 
            this.addOne.Label = "確認";
            this.addOne.Name = "addOne";
            this.addOne.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.addOne_Click);
            // 
            // addSTVmode
            // 
            this.addSTVmode.Items.Add(this.sourceSelectMode);
            this.addSTVmode.Items.Add(this.targetSelectMode);
            this.addSTVmode.Items.Add(this.valueSelectMode);
            this.addSTVmode.Label = "指標";
            this.addSTVmode.Name = "addSTVmode";
            // 
            // sourceSelectMode
            // 
            this.sourceSelectMode.Checked = true;
            this.sourceSelectMode.Label = "選擇「來源」";
            this.sourceSelectMode.Name = "sourceSelectMode";
            this.sourceSelectMode.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.sourceSelectMode_Click);
            // 
            // targetSelectMode
            // 
            this.targetSelectMode.Label = "選擇「目標」";
            this.targetSelectMode.Name = "targetSelectMode";
            this.targetSelectMode.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.targetSelectMode_Click);
            // 
            // valueSelectMode
            // 
            this.valueSelectMode.Label = "選擇「值」";
            this.valueSelectMode.Name = "valueSelectMode";
            this.valueSelectMode.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.valueSelectMode_Click);
            // 
            // removeSelection
            // 
            ribbonDropDownItemImpl1.Label = "來源(尚未選取)";
            ribbonDropDownItemImpl2.Label = "目標(尚未選取)";
            ribbonDropDownItemImpl3.Label = "值(尚未選取)";
            this.removeSelection.Items.Add(ribbonDropDownItemImpl1);
            this.removeSelection.Items.Add(ribbonDropDownItemImpl2);
            this.removeSelection.Items.Add(ribbonDropDownItemImpl3);
            this.removeSelection.Label = "檢視";
            this.removeSelection.Name = "removeSelection";
            this.removeSelection.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.removeSelection_SelectionChanged);
            // 
            // doRemoveSelection
            // 
            this.doRemoveSelection.Label = "移除";
            this.doRemoveSelection.Name = "doRemoveSelection";
            this.doRemoveSelection.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.doRemoveSelection_Click);
            // 
            // settingDC
            // 
            this.settingDC.Items.Add(this.autoNextPT);
            this.settingDC.Items.Add(this.autoPreView);
            this.settingDC.Items.Add(this.hotfixAutoReset41);
            this.settingDC.Items.Add(this.safe3);
            this.settingDC.Label = "指標自動化設定";
            this.settingDC.Name = "settingDC";
            // 
            // autoNextPT
            // 
            this.autoNextPT.Checked = true;
            this.autoNextPT.Label = "指標自遞增";
            this.autoNextPT.Name = "autoNextPT";
            // 
            // autoPreView
            // 
            this.autoPreView.Checked = true;
            this.autoPreView.Label = "自動預檢";
            this.autoPreView.Name = "autoPreView";
            // 
            // hotfixAutoReset41
            // 
            this.hotfixAutoReset41.Checked = true;
            this.hotfixAutoReset41.Label = "自動復歸";
            this.hotfixAutoReset41.Name = "hotfixAutoReset41";
            // 
            // safe3
            // 
            this.safe3.Checked = true;
            this.safe3.Label = "捕捉缺項錯誤";
            this.safe3.Name = "safe3";
            // 
            // kpOperate
            // 
            this.kpOperate.Items.Add(this.arraySetSource);
            this.kpOperate.Items.Add(this.arraySetTarget);
            this.kpOperate.Items.Add(this.previewArray);
            this.kpOperate.Items.Add(this.arraySetData);
            this.kpOperate.Items.Add(this.arrayColorSetting);
            this.kpOperate.Label = "kpOperate";
            this.kpOperate.Name = "kpOperate";
            // 
            // arraySetSource
            // 
            this.arraySetSource.Label = "arraySetSource";
            this.arraySetSource.Name = "arraySetSource";
            this.arraySetSource.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.arraySetSource_Click);
            // 
            // arraySetTarget
            // 
            this.arraySetTarget.Label = "arraySetTarget";
            this.arraySetTarget.Name = "arraySetTarget";
            this.arraySetTarget.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.arraySetTarget_Click);
            // 
            // previewArray
            // 
            this.previewArray.Label = "previewArray";
            this.previewArray.Name = "previewArray";
            this.previewArray.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.previewArray_Click);
            // 
            // arraySetData
            // 
            this.arraySetData.Label = "arraySetData";
            this.arraySetData.Name = "arraySetData";
            this.arraySetData.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.arraySetData_Click);
            // 
            // arrayColorSetting
            // 
            this.arrayColorSetting.Items.Add(this.displayColorAfterSelect);
            this.arrayColorSetting.Items.Add(this.picColor1);
            this.arrayColorSetting.Items.Add(this.picColor2);
            this.arrayColorSetting.Items.Add(this.picColor3);
            this.arrayColorSetting.Items.Add(this.picColor4);
            this.arrayColorSetting.Items.Add(this.picColor5);
            this.arrayColorSetting.Items.Add(this.picColor6);
            this.arrayColorSetting.Label = "arrayColorSetting";
            this.arrayColorSetting.Name = "arrayColorSetting";
            this.arrayColorSetting.ShowImage = true;
            // 
            // displayColorAfterSelect
            // 
            this.displayColorAfterSelect.Checked = true;
            this.displayColorAfterSelect.Label = "displayColorAfterSelect";
            this.displayColorAfterSelect.Name = "displayColorAfterSelect";
            // 
            // picColor1
            // 
            this.picColor1.Label = "picColor1";
            this.picColor1.Name = "picColor1";
            this.picColor1.ShowImage = true;
            this.picColor1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.picColor1_Click);
            // 
            // picColor2
            // 
            this.picColor2.Label = "picColor2";
            this.picColor2.Name = "picColor2";
            this.picColor2.ShowImage = true;
            this.picColor2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.picColor2_Click);
            // 
            // picColor3
            // 
            this.picColor3.Label = "picColor3";
            this.picColor3.Name = "picColor3";
            this.picColor3.ShowImage = true;
            this.picColor3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.picColor3_Click);
            // 
            // picColor4
            // 
            this.picColor4.Label = "picColor4";
            this.picColor4.Name = "picColor4";
            this.picColor4.ShowImage = true;
            this.picColor4.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.picColor4_Click);
            // 
            // picColor5
            // 
            this.picColor5.Label = "picColor5";
            this.picColor5.Name = "picColor5";
            this.picColor5.ShowImage = true;
            this.picColor5.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.picColor5_Click);
            // 
            // picColor6
            // 
            this.picColor6.Label = "picColor6";
            this.picColor6.Name = "picColor6";
            this.picColor6.ShowImage = true;
            this.picColor6.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.picColor6_Click);
            // 
            // primaryOP
            // 
            this.primaryOP.Label = "primaryOP";
            this.primaryOP.Name = "primaryOP";
            // 
            // juniorOP
            // 
            this.juniorOP.Items.Add(this.addSplitButton);
            this.juniorOP.Items.Add(this.testAddRow);
            this.juniorOP.Label = "juniorOP";
            this.juniorOP.Name = "juniorOP";
            // 
            // addSplitButton
            // 
            this.addSplitButton.Label = "開啟遮罩視窗";
            this.addSplitButton.Name = "addSplitButton";
            this.addSplitButton.ShowImage = true;
            this.addSplitButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.addSplitButton_Click);
            // 
            // testAddRow
            // 
            this.testAddRow.Label = "加入遮罩";
            this.testAddRow.Name = "testAddRow";
            this.testAddRow.ShowImage = true;
            this.testAddRow.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.testAddRow_Click);
            // 
            // seniorOP
            // 
            this.seniorOP.Label = "seniorOP";
            this.seniorOP.Name = "seniorOP";
            // 
            // mgOperate
            // 
            this.mgOperate.Items.Add(this.modeEdit);
            this.mgOperate.Items.Add(this.editData);
            this.mgOperate.Items.Add(this.clearVisual);
            this.mgOperate.Items.Add(this.mgOperateColorSetting);
            this.mgOperate.Items.Add(this.rainbowMG);
            this.mgOperate.Label = "mgOperate";
            this.mgOperate.Name = "mgOperate";
            // 
            // modeEdit
            // 
            ribbonDropDownItemImpl4.Label = "檢視";
            ribbonDropDownItemImpl4.Tag = "view";
            ribbonDropDownItemImpl5.Label = "移除";
            ribbonDropDownItemImpl5.Tag = "del";
            this.modeEdit.Items.Add(ribbonDropDownItemImpl4);
            this.modeEdit.Items.Add(ribbonDropDownItemImpl5);
            this.modeEdit.Label = "modeEdit";
            this.modeEdit.Name = "modeEdit";
            this.modeEdit.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.modeEdit_SelectionChanged);
            // 
            // editData
            // 
            ribbonDropDownItemImpl6.Label = "Item0";
            this.editData.Items.Add(ribbonDropDownItemImpl6);
            this.editData.Label = "編輯資料";
            this.editData.Name = "editData";
            this.editData.ScreenTip = "滑鼠懸停在項目上方可以檢視其儲存格的位置";
            this.editData.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.editData_Click);
            this.editData.ItemsLoading += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.editDataLoad);
            // 
            // clearVisual
            // 
            this.clearVisual.Label = "clearVisual";
            this.clearVisual.Name = "clearVisual";
            this.clearVisual.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.clearVisual_Click);
            // 
            // mgOperateColorSetting
            // 
            this.mgOperateColorSetting.Items.Add(this.a1show);
            this.mgOperateColorSetting.Items.Add(this.a2show);
            this.mgOperateColorSetting.Items.Add(this.b1show);
            this.mgOperateColorSetting.Items.Add(this.b2show);
            this.mgOperateColorSetting.Items.Add(this.c1show);
            this.mgOperateColorSetting.Items.Add(this.c2show);
            this.mgOperateColorSetting.Label = "mgOperateColorSetting";
            this.mgOperateColorSetting.Name = "mgOperateColorSetting";
            // 
            // a1show
            // 
            this.a1show.Label = "來源-背景";
            this.a1show.Name = "a1show";
            this.a1show.ShowImage = true;
            this.a1show.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.a1show_Click);
            // 
            // a2show
            // 
            this.a2show.Label = "來源-文字";
            this.a2show.Name = "a2show";
            this.a2show.ShowImage = true;
            this.a2show.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.a2show_Click);
            // 
            // b1show
            // 
            this.b1show.Label = "目標-背景";
            this.b1show.Name = "b1show";
            this.b1show.ShowImage = true;
            this.b1show.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.b1show_Click);
            // 
            // b2show
            // 
            this.b2show.Label = "目標-文字";
            this.b2show.Name = "b2show";
            this.b2show.ShowImage = true;
            this.b2show.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.b2show_Click);
            // 
            // c1show
            // 
            this.c1show.Label = "值-背景";
            this.c1show.Name = "c1show";
            this.c1show.ShowImage = true;
            this.c1show.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.c1show_Click);
            // 
            // c2show
            // 
            this.c2show.Label = "值-文字";
            this.c2show.Name = "c2show";
            this.c2show.ShowImage = true;
            this.c2show.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.c2show_Click);
            // 
            // rainbowMG
            // 
            this.rainbowMG.Label = "資料校對動畫";
            this.rainbowMG.Name = "rainbowMG";
            this.rainbowMG.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.rainbowMG_Click);
            // 
            // dpOperate
            // 
            this.dpOperate.Items.Add(this.displayData);
            this.dpOperate.Items.Add(this.processData);
            this.dpOperate.Items.Add(this.plotTitle);
            this.dpOperate.Items.Add(this.autoNodeColorSetting);
            this.dpOperate.Label = "dpOperate";
            this.dpOperate.Name = "dpOperate";
            // 
            // displayData
            // 
            this.displayData.Label = "displayData";
            this.displayData.Name = "displayData";
            this.displayData.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.displayData_Click);
            // 
            // processData
            // 
            this.processData.Label = "processData";
            this.processData.Name = "processData";
            this.processData.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.processData_Click);
            // 
            // plotTitle
            // 
            this.plotTitle.Label = "圖主名稱";
            this.plotTitle.Name = "plotTitle";
            this.plotTitle.Text = null;
            // 
            // autoNodeColorSetting
            // 
            this.autoNodeColorSetting.Checked = true;
            this.autoNodeColorSetting.Label = "自動生成節點顏色";
            this.autoNodeColorSetting.Name = "autoNodeColorSetting";
            // 
            // ucOperate
            // 
            this.ucOperate.Items.Add(this.saveMap);
            this.ucOperate.Items.Add(this.loadMap);
            this.ucOperate.Items.Add(this.exportMap);
            this.ucOperate.Items.Add(this.importMap);
            this.ucOperate.Items.Add(this.setting4save);
            this.ucOperate.Items.Add(this.shiftSetting);
            this.ucOperate.Label = "ucOperate";
            this.ucOperate.Name = "ucOperate";
            // 
            // saveMap
            // 
            this.saveMap.Label = "存檔";
            this.saveMap.Name = "saveMap";
            this.saveMap.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.saveMap_Click);
            // 
            // loadMap
            // 
            this.loadMap.Label = "載入";
            this.loadMap.Name = "loadMap";
            this.loadMap.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.loadMap_Click);
            // 
            // exportMap
            // 
            this.exportMap.Label = "匯出";
            this.exportMap.Name = "exportMap";
            this.exportMap.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.exportMap_Click);
            // 
            // importMap
            // 
            this.importMap.Label = "匯入";
            this.importMap.Name = "importMap";
            this.importMap.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.importMap_Click);
            // 
            // setting4save
            // 
            this.setting4save.Items.Add(this.emptyWhenLoad);
            this.setting4save.Items.Add(this.ableShift);
            this.setting4save.Label = "IO設定";
            this.setting4save.Name = "setting4save";
            // 
            // emptyWhenLoad
            // 
            this.emptyWhenLoad.Checked = true;
            this.emptyWhenLoad.Label = "載入時清空主記憶體(未打勾時執行複寫操作)";
            this.emptyWhenLoad.Name = "emptyWhenLoad";
            // 
            // ableShift
            // 
            this.ableShift.Checked = true;
            this.ableShift.Label = "ableShift";
            this.ableShift.Name = "ableShift";
            // 
            // shiftSetting
            // 
            this.shiftSetting.Label = "shiftSetting";
            this.shiftSetting.Name = "shiftSetting";
            this.shiftSetting.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.shiftSetting_Click);
            this.shiftSetting.ItemsLoading += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.updateWorkSheetShiftSetting);
            // 
            // pdOperate
            // 
            this.pdOperate.Items.Add(this.debugHide);
            this.pdOperate.Items.Add(this.Rpath);
            this.pdOperate.Items.Add(this.musicPath);
            this.pdOperate.Items.Add(this.useOldR);
            this.pdOperate.Label = "pdOperate";
            this.pdOperate.Name = "pdOperate";
            // 
            // debugHide
            // 
            this.debugHide.Items.Add(this.addMainData);
            this.debugHide.Items.Add(this.readUserSelectB);
            this.debugHide.Items.Add(this.addRibbonDropdownItemB);
            this.debugHide.Items.Add(this.listTest);
            this.debugHide.Items.Add(this.writeMainDataDumb);
            this.debugHide.Items.Add(this.galleryNumTest);
            this.debugHide.Items.Add(this.rainbowTest);
            this.debugHide.Items.Add(this.testsave);
            this.debugHide.Items.Add(this.testloadsave);
            this.debugHide.Items.Add(this.worksheetcodenametest);
            this.debugHide.Items.Add(this.testActivateWindows);
            this.debugHide.Items.Add(this.newWindowsTag);
            this.debugHide.Items.Add(this.assembHTML);
            this.debugHide.Items.Add(this.defaultSnakeColorTest);
            this.debugHide.Items.Add(this.todolist);
            this.debugHide.Label = "debugHide";
            this.debugHide.Name = "debugHide";
            // 
            // addMainData
            // 
            this.addMainData.Label = "addMainData";
            this.addMainData.Name = "addMainData";
            this.addMainData.ShowImage = true;
            this.addMainData.SuperTip = "選3個cell，按序:輸入，輸出，值";
            this.addMainData.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.addMainData_Click);
            // 
            // readUserSelectB
            // 
            this.readUserSelectB.Label = "readUserSelect";
            this.readUserSelectB.Name = "readUserSelectB";
            this.readUserSelectB.ShowImage = true;
            this.readUserSelectB.SuperTip = "選取讀取測試";
            this.readUserSelectB.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.readUserSelectB_Click);
            // 
            // addRibbonDropdownItemB
            // 
            this.addRibbonDropdownItemB.Label = "外部執行測試";
            this.addRibbonDropdownItemB.Name = "addRibbonDropdownItemB";
            this.addRibbonDropdownItemB.ShowImage = true;
            this.addRibbonDropdownItemB.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.addRibbonDropdownItemB_Click);
            // 
            // listTest
            // 
            this.listTest.Label = "listTest";
            this.listTest.Name = "listTest";
            this.listTest.ShowImage = true;
            this.listTest.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.listTest_Click);
            // 
            // writeMainDataDumb
            // 
            this.writeMainDataDumb.Label = "writeMainDataDumb";
            this.writeMainDataDumb.Name = "writeMainDataDumb";
            this.writeMainDataDumb.ShowImage = true;
            this.writeMainDataDumb.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.writeMainDataDumb_Click);
            // 
            // galleryNumTest
            // 
            this.galleryNumTest.Label = "galleryNumTest";
            this.galleryNumTest.Name = "galleryNumTest";
            this.galleryNumTest.ShowImage = true;
            this.galleryNumTest.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.galleryNumTest_Click);
            // 
            // rainbowTest
            // 
            this.rainbowTest.Label = "ranbowTest";
            this.rainbowTest.Name = "rainbowTest";
            this.rainbowTest.ShowImage = true;
            this.rainbowTest.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.rainbowTest_Click);
            // 
            // testsave
            // 
            this.testsave.Label = "testsave";
            this.testsave.Name = "testsave";
            this.testsave.ShowImage = true;
            this.testsave.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.testsave_Click);
            // 
            // testloadsave
            // 
            this.testloadsave.Label = "testloadsave";
            this.testloadsave.Name = "testloadsave";
            this.testloadsave.ShowImage = true;
            this.testloadsave.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.testloadsave_Click);
            // 
            // worksheetcodenametest
            // 
            this.worksheetcodenametest.Label = "worksheetcodenametest";
            this.worksheetcodenametest.Name = "worksheetcodenametest";
            this.worksheetcodenametest.ShowImage = true;
            this.worksheetcodenametest.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.worksheetcodenametest_Click);
            // 
            // testActivateWindows
            // 
            this.testActivateWindows.Label = "testActivateWindows";
            this.testActivateWindows.Name = "testActivateWindows";
            this.testActivateWindows.ShowImage = true;
            this.testActivateWindows.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.testActivateWindows_Click);
            // 
            // newWindowsTag
            // 
            this.newWindowsTag.Label = "newWindowsTag";
            this.newWindowsTag.Name = "newWindowsTag";
            this.newWindowsTag.ShowImage = true;
            this.newWindowsTag.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.newWindowsTag_Click);
            // 
            // defaultSnakeColorTest
            // 
            this.defaultSnakeColorTest.Label = "defaultSnakeColorTest";
            this.defaultSnakeColorTest.Name = "defaultSnakeColorTest";
            this.defaultSnakeColorTest.ShowImage = true;
            this.defaultSnakeColorTest.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.defaultSnakeColorTest_Click);
            // 
            // todolist
            // 
            this.todolist.Label = "todolist";
            this.todolist.Name = "todolist";
            this.todolist.ShowImage = true;
            this.todolist.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.todolist_Click);
            // 
            // Rpath
            // 
            this.Rpath.Label = "Rpath";
            this.Rpath.Name = "Rpath";
            this.Rpath.Text = "C:\\Users\\ai\\Documents\\andy\\code\\snakeskin\\masterR";
            // 
            // musicPath
            // 
            this.musicPath.Label = "musicPath";
            this.musicPath.Name = "musicPath";
            this.musicPath.Text = "C:\\Users\\ai\\Music\\akbS63.wav";
            // 
            // useOldR
            // 
            this.useOldR.Label = "使用舊版渲染R程式";
            this.useOldR.Name = "useOldR";
            this.useOldR.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.useOldR_Click);
            // 
            // a1
            // 
            this.a1.Color = System.Drawing.SystemColors.Window;
            // 
            // saveMirrorText
            // 
            this.saveMirrorText.DefaultExt = "json";
            this.saveMirrorText.FileName = "蛇圖格關聯檔";
            this.saveMirrorText.Filter = "\"JSON files (*.json)|*.json|All files (*.*)|*.*\"";
            this.saveMirrorText.RestoreDirectory = true;
            this.saveMirrorText.Title = "正在匯出蛇圖格關聯檔";
            // 
            // loadMirrorText
            // 
            this.loadMirrorText.DefaultExt = "json";
            this.loadMirrorText.FileName = "蛇圖格關聯檔";
            this.loadMirrorText.Filter = "\"JSON files (*.json)|*.json|All files (*.*)|*.*\"";
            this.loadMirrorText.RestoreDirectory = true;
            this.loadMirrorText.Title = "正在匯入蛇圖格關聯檔";
            // 
            // assembHTML
            // 
            this.assembHTML.Label = "assemblyHTML";
            this.assembHTML.Name = "assembHTML";
            this.assembHTML.ShowImage = true;
            this.assembHTML.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.assembHTML_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.dcOperate.ResumeLayout(false);
            this.dcOperate.PerformLayout();
            this.kpOperate.ResumeLayout(false);
            this.kpOperate.PerformLayout();
            this.juniorOP.ResumeLayout(false);
            this.juniorOP.PerformLayout();
            this.mgOperate.ResumeLayout(false);
            this.mgOperate.PerformLayout();
            this.dpOperate.ResumeLayout(false);
            this.dpOperate.PerformLayout();
            this.ucOperate.ResumeLayout(false);
            this.ucOperate.PerformLayout();
            this.pdOperate.ResumeLayout(false);
            this.pdOperate.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup dcOperate;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup kpOperate;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup primaryOP;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup juniorOP;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup seniorOP;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup ucOperate;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup mgOperate;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup dpOperate;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup pdOperate;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton displayData;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton addMainData;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton capture;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown removeSelection;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton addOne;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu addSTVmode;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox sourceSelectMode;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox targetSelectMode;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox valueSelectMode;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton doRemoveSelection;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton readUserSelectB;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton addRibbonDropdownItemB;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu settingDC;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox autoNextPT;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox autoPreView;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton processData;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu debugHide;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton listTest;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton todolist;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton writeMainDataDumb;
        internal Microsoft.Office.Tools.Ribbon.RibbonGallery editData;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton galleryNumTest;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown modeEdit;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton clearVisual;
        private System.Windows.Forms.ColorDialog a1;
        private System.Windows.Forms.ColorDialog a2;
        private System.Windows.Forms.ColorDialog b1;
        private System.Windows.Forms.ColorDialog b2;
        private System.Windows.Forms.ColorDialog c1;
        private System.Windows.Forms.ColorDialog c2;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu mgOperateColorSetting;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton a1show;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton a2show;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton b1show;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton b2show;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton c1show;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton c2show;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton arraySetSource;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton arraySetTarget;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton arraySetData;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton previewArray;
        private System.Windows.Forms.ColorDialog arrayColorSetSource1;
        private System.Windows.Forms.ColorDialog arrayColorSetSource2;
        private System.Windows.Forms.ColorDialog arrayColorSetTarget1;
        private System.Windows.Forms.ColorDialog arrayColorSetTarget2;
        private System.Windows.Forms.ColorDialog arrayColorSetData1;
        private System.Windows.Forms.ColorDialog arrayColorSetData2;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu arrayColorSetting;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox displayColorAfterSelect;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton picColor1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton picColor2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton picColor3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton picColor4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton picColor5;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton picColor6;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox hotfixAutoReset41;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox safe3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton rainbowTest;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton rainbowMG;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton testsave;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton testloadsave;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton worksheetcodenametest;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton saveMap;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton loadMap;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton exportMap;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton importMap;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu setting4save;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox emptyWhenLoad;
        private System.Windows.Forms.SaveFileDialog saveMirrorText;
        private System.Windows.Forms.OpenFileDialog loadMirrorText;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox ableShift;
        internal Microsoft.Office.Tools.Ribbon.RibbonGallery shiftSetting;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton testActivateWindows;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton addSplitButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton testAddRow;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton newWindowsTag;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox Rpath;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox musicPath;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton useOldR;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton defaultSnakeColorTest;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox plotTitle;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton autoNodeColorSetting;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton assembHTML;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
