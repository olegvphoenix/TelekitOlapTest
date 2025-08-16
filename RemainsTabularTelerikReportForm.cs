using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using DevExpress.Utils;
using DevExpress.XtraBars;
using DevExpress.XtraCharts;
using DevExpress.XtraEditors;
using DevExpress.XtraEditors.Repository;
using DevExpress.XtraPivotGrid;
using DevExpress.XtraPrinting;
using DevExpress.XtraTab;

using TruckService.ADTS.Client.BaseUI.AppForms.BaseForms;
using TruckService.ADTS.Client.BaseUI.AppForms.Modal;
using TruckService.ADTS.Client.BaseUI.Utility;
using TruckService.ADTS.Client.BaseUI.Utility.ExcelExport;
using TruckService.ADTS.Client.Data.GridFormSettings.FilterControls;
using TruckService.ADTS.Client.Utility;

using TruckService.ADTS.Contractor.Utility;
using TruckService.ADTS.Controls;
using TruckService.ADTS.Controls.Accessors;
using TruckService.ADTS.Controls.Enums;
using TruckService.ADTS.Controls.Utility;
using TruckService.ADTS.Core.DAL;
using TruckService.ADTS.Core.DAL.Entities;
using TruckService.ADTS.Core.Lists;
using TruckService.ADTS.Core.Utility;
using PaymentType = TruckService.ADTS.Core.DAL.Entities.PaymentType;

namespace TruckService.ADTS.Client.AppForms.Grid
{
    /// <summary>
    /// Комплексный отчёт по товародвижению (табулярная модель, Остатки).
    /// </summary>
    public partial class RemainsTabularTelerikReportForm : BaseChildForm
    {
        private const string GridLayoutSettingsCode = "RemainsTabular_Report_Pivot_Grid_Code";
        private const string GridChartLayoutSettingsCode = "RemainsTabular_Report_Pivot_Grid_Chart_Code";

        #region Fields

        private bool _isGridInitialized = false;
        private bool _isFormInitialized = false;
        private bool _isChartInitialized = false;
        private bool _isLegendVisible = true;
        private bool _isPointsLabelVisible = true;
        private WaitCursor _waitCursor = null;
        private PaymentTypeExceptComissionMultiSelectControl _paymentTypeFilter;
        private readonly ContractorSelectHelper _contractorSelectHelper = new ContractorSelectHelper();
        private readonly string _windowGuid;
        private string _sessionGuid;
        private readonly ApplicationSettingsRepository _settings;

        // последнее выставленное представление
        private string _presentation = null;
        private BarItemLinkCollection _menuButtons;
        private BarItemLinkCollection _buttons;
        private readonly List<ToolBarButtonIndex> _attachedButtons = new List<ToolBarButtonIndex>();

        private object _lastGridSettings = null;
        private readonly object _lastGridChartSettings = null;

        private bool _isFilterAppliedToChart = false;
        private bool _isFilterAppliedToGrid = false;
        private bool _isAsyncInProgress = false;

        private readonly StorageSelectHelper _storageSelectHelper = new StorageSelectHelper();

        private readonly string[] ControlFiltersNamesForSheetHeaders =
        {
            nameof(txtPeriodFrom),
            nameof(txtPeriodTo),
            nameof(cmbProducer),
            nameof(cmbAssortimentGoods),
            nameof(txtContract),
            nameof(txtAccepStorage),
            nameof(txtResponsibleManager),
            nameof(cmbBasisOfPurchase),
            "cmbPaymentType",
            nameof(cmbOrderType)
        };

        private const string PERIOD_FIELD_NAME = "[Дата-Календарь].[День].[День]";

        // список полей по которым выполняется ручная сортировка в ServerMode по MEMBER_KEY
        private readonly string[] FieldsForCustomSort =
        {
            PERIOD_FIELD_NAME
        };

        // список составных полей по которым выполняется выполняется ручная сортировка в ServerMode
        private readonly string[] CompositeFieldsForCustomSort =
        {
            "[Дата-Календарь].[Неделя].[Неделя]",
            "[Дата-Календарь].[МесяцКод].[МесяцКод]",
            "[Дата-Календарь].[Месяц].[Месяц]",
            "[Дата-Календарь].[Квартал].[Квартал]"
        };

        #endregion

        #region Properties

        public override bool CanMdiMultiInstance => true;


        private bool IsAnyFilterParamSet
        {
            get
            {
                foreach (Control c in pnlFilter.Controls)
                {
                    if (c is BaseEdit edit && !edit.EditValue.IsNullOrDBNull())
                        return true;
                }

                return false;
            }
        }

        private string DateFilter => "[Дата-Календарь].[Дата].[Дата]";

        #endregion

        #region Inner classes
        //описание панели фильтров
        public class RemainsFilterValues
        {
            public object PeriodFrom;
            public object PeriodTo;

            public string SkubaNumber;    //СКН номер детали
            public string PartName;       //название детали
            public string UniqueNumber;   //уникальный номер детали
            public string PartCode;       //код детали
            public string GoodsCode;      //код товара 
            public object GroupList;      // группы
            public object SubgroupList;   // подгруппы
            public object ProducerList;   // производитель

            public object ContractorIdList; // поставщик
            public object StorageIdList;  // склад
            public object AbcList;        // ABC детали
            public object AbcGoodsList;   // ABC товара
            public object XyzList;        // XYZ товара
            public object SalesOperation; // операция продажи
            public object ManagerIdList;  // ответственный
            public object PaymentType;    // тип платежа
            public object AssortimentGoodsList; // ассортиментный товар
        }

        #endregion

        #region Constructors

        public RemainsTabularTelerikReportForm()
            : base(null)
        {
            InitializeComponent();

            _settings = Metadata.RepositoryFactory.Get<ApplicationSettings, ApplicationSettingsRepository>(Database);

            _windowGuid = OlapUtility.GetNextSessionId();
            _sessionGuid = OlapUtility.GetNextSessionId();
        }

        #endregion

        private void InitGrid()
        {
            if (_isGridInitialized)
                return;
            _isGridInitialized = true;

            pivotGrid.CellSelectionChanged += pivotGrid_CellSelectionChanged;
            pivotGrid.CellClick += pivotGrid_CellClick;
            pivotGrid.FieldValueExpanded += pivotGrid_FieldValueExpanded;
            pivotGrid.FieldValueCollapsed += pivotGrid_FieldValueCollapsed;
            pivotGrid.FieldValueNotExpanded += pivotGrid_FieldValueNotExpanded;
            pivotGrid.FieldValueExpanding += pivotGrid_FieldValueExpanding;

            pivotGridControlForChart.CellSelectionChanged += pivotGridControlForChart_CellSelectionChanged;
            pivotGridControlForChart.CellClick += pivotGridControlForChart_CellClick;
            pivotGridControlForChart.FieldValueExpanded += pivotGridControlForChart_FieldValueExpanded;
            pivotGridControlForChart.FieldValueCollapsed += pivotGridControlForChart_FieldValueCollapsed;
        }

        private void ConfigureAndPopulateLookUps()
        {
            _contractorSelectHelper.RegisterEdit(ContractorSelectHelper.Descriptor.GetContractor(txtContract, showResetButton: true, isMultiSelect: true, showRetailCheckBox: true));
            txtContract.Properties.NullText = DataUtility.GetAnyText(AnyTextType.Empty);

            ListControlUtility.PopulateGoodsSaleRatingCombo(Database, cmbGoodSaleRating);
            ListControlUtility.PopulatePartSaleRatingCombo(Database, cmbPartSaleRating);
            ListControlUtility.PopulateGoodsProfitRatingCombo(cmbGoodsProfitRating);

            ListControlUtility.PopulateCostCheckedCombo(cmbXyz);

            ListControlUtility.PopulateGoodsSpecifications(cmbGoodsSpecifications);
            ListControlUtility.PopulatePartSpecifications(cmbPartSpecifications);

            ListControlUtility.PopulateBasisOfPurchase(cmbBasisOfPurchase);

            ListControlUtility.PopulateOrderType(cmbOrderType);

            ListControlUtility.PopulateGroupCheckedCombo(cmbGroup);
            var producerSelectHelper = new ProducerSelectHelper(); //#490697
            producerSelectHelper.RegisterEdit(cmbProducer);

            var ctSubGroupSelectHelper = new SubGroupMultiSelectHelper();
            ctSubGroupSelectHelper.RegisterEdit(cmbSubgroup);

            PopulateAssortmentGoodsCheckedCombo(cmbAssortimentGoods);

            _storageSelectHelper.RegisterEdit(txtAccepStorage, isExternal: false, isMultiSelect: true, showResetButton: true, selectStorageGroupSegments: true);

            _paymentTypeFilter = new PaymentTypeExceptComissionMultiSelectControl();
            var cmbPaymentType = (CheckedComboBoxEdit)_paymentTypeFilter.CreateFilterControl();
            cmbPaymentType.Location = new Point(lblPaymentType.Location.X, 71);
            cmbPaymentType.Name = "cmbPaymentType";
            cmbPaymentType.Size = new Size(130, 20);
            cmbPaymentType.TabIndex = 41;
            pnlFilter.Controls.Add(cmbPaymentType);
            _paymentTypeFilter.EditValue = $"{PaymentType.GetPaymentTypeId(Core.Lists.PaymentType.Cashless)}, {PaymentType.GetPaymentTypeId(Core.Lists.PaymentType.OnHand)}";

            cmbOperation.Properties.NullText = DataUtility.GetAnyText(AnyTextType.Empty);
            LookupIntLists<SalesOperation>.PopulateLookup(cmbOperation, true, true, false);

            _contractorSelectHelper.RegisterEdit(ContractorSelectHelper.Descriptor.GetResponsibleManager(txtResponsibleManager));
            txtResponsibleManager.Properties.NullText = DataUtility.GetAnyText(AnyTextType.Empty);

            _contractorSelectHelper.RegisterEdit(ContractorSelectHelper.Descriptor.GetResponsibleManager(txtResponsibleForDeliveryToBranch));
            txtResponsibleForDeliveryToBranch.Properties.NullText = DataUtility.GetAnyText(AnyTextType.Empty);

            txtUniqueNumber.TopLimitCount
                = txtSKNNumber.TopLimitCount
                = txtGoodsCode.TopLimitCount
                = txtPartCode.TopLimitCount = SystemGlobalSettings.GetMaxInexactSearchNumberCount(Database);
        }

        private void StartWaitCursor()
        {
            if (_waitCursor == null)
                _waitCursor = new WaitCursor();
        }

        private void EndWaitCursor()
        {
            if (_waitCursor != null)
            {
                _waitCursor.Dispose();
                _waitCursor = null;
            }
        }

        //скинуть все установленные значения фильтров, период установить по умолчанибю прошлый месяц до вчерашнего дня
        private void ResetFilter()
        {
            foreach (var c in pnlFilter.Controls)
            {
                if (c is BaseEdit edit)
                    edit.EditValue = DBNull.Value;
            }

            txtSKNNumber.EditValue = null;
            txtGoodsCode.EditValue = null;
            txtPartCode.EditValue = null;
            txtUniqueNumber.EditValue = null;

            txtPeriodTo.EditValue = DateTime.Today.AddDays(-1).Date;
            txtPeriodFrom.EditValue = DateTime.Today.AddMonths(-1).Date;
            cmbAssortimentGoods.CheckAll();

            if (tabChilds.SelectedTabPage == tbpChart)
                BindChart();
            else if (tabChilds.SelectedTabPage == tbpGrid)
                BindData();
        }

        //применить установленные значения фильтров к фильтрам активного pivottable
        private void ApplyFilter()
        {
            try
            {
                loadManager.StartTimer();

                InitialiseCube();

                _isFilterAppliedToChart = false;
                _isFilterAppliedToGrid = false;
                if (tabChilds.SelectedTabPage == tbpChart)
                {
                    if (!_isFilterAppliedToChart)
                    {
                        pivotGridControlForChart.BeginUpdate();
                        try
                        {
                            BindChart();
                        }
                        finally
                        {
                            pivotGridControlForChart.EndUpdate();
                        }
                        _isFilterAppliedToChart = true;
                    }
                }
                else if (tabChilds.SelectedTabPage == tbpGrid)
                {
                    if (!_isFilterAppliedToGrid && !pivotGrid.IsAsyncInProgress)
                    {
                        pivotGrid.BeginUpdate();
                        try
                        {
                            BindData();
                        }
                        finally
                        {
                            pivotGrid.EndUpdateAsync();
                        }

                        _isFilterAppliedToGrid = true;
                    }
                }
            }
            finally
            {
                loadManager.StopTimer();
            }
        }

        private void BindData()
        {
            if (pivotGrid.IsAsyncInProgress)
                return;
            
            TrimTimeFromDateFields();
            if (IsAnyFilterParamSet)
                TraceSearch();

            OlapUtility.SetPivotFieldDateFilter(pivotGrid, DateFilter, txtPeriodFrom.EditValue, txtPeriodTo.EditValue);

            // Группа cmbGroup - id_group - 
            OlapUtility.SetPivotFieldIdFilter(pivotGrid, "[Номенклатура-Группа].[Ключ].[Ключ]", FilterUtility.GetFilterKeyValue(cmbGroup.EditValue), true);

            // Подгруппа  cmbSubgroup - id_subgroup - 
            var subgroupFilterValue = FilterUtility.GetSubgroupFilter(cmbSubgroup.EditValue);
            OlapUtility.SetPivotFieldIdFilter(pivotGrid, "[Номенклатура-Подгруппа].[Код подгруппы].[Код подгруппы]", subgroupFilterValue, true);

            // Склад   txtStorage - storages_list - 
            var ShpStorageFilterValue = FilterUtility.GetStorageFilter(txtAccepStorage.EditValue);
            OlapUtility.SetPivotFieldIdFilter(pivotGrid, "[Склад].[Ключ абстрактный].[Ключ абстрактный]", ShpStorageFilterValue, true);

            // Тип операции
            // cmbPaymentType - payment_type_list - [Тип операции].[Ключ].[Ключ]
            var paymentTypeFilter = FilterUtility.GetFilterKeyValue(_paymentTypeFilter.Control.EditValue);
            if (paymentTypeFilter == null)
                paymentTypeFilter = new object[0];

            OlapUtility.SetPivotFieldIdFilter(pivotGrid, "[Тип операции].[Ключ].[Ключ]", paymentTypeFilter, true);

            OlapUtility.SetPivotFieldIdFilter(pivotGrid, "[Поставщик].[Ключ].[Ключ]", FilterUtility.GetFilterKeyValue(txtContract.EditValue), true);

            // cmbProducer 
            var producerFilterValue = FilterUtility.GetProducerFilter(cmbProducer.EditValue);
            OlapUtility.SetPivotFieldIdFilter(pivotGrid, "[Номенклатура].[ПроизводительКлюч].[ПроизводительКлюч]", producerFilterValue, true);

            OlapUtility.SetPivotFieldIdFilter(pivotGrid, "[Ответственный за подгруппу].[Id_abstract_contractor].[Id_abstract_contractor]", FilterUtility.GetFilterKeyValue(txtResponsibleManager.EditValue), true);

            OlapUtility.SetPivotFieldIdFilter(pivotGrid, "[Ответственный за филиал].[id_responsible_abstract].[id_responsible_abstract]", FilterUtility.GetFilterKeyValue(txtResponsibleForDeliveryToBranch.EditValue), true);

            OlapUtility.SetPivotFieldIdFilter(pivotGrid, "[Рейтинг по прибыли].[ABC по рейтингу товара по прибыли].[ABC по рейтингу товара по прибыли]", FilterUtility.GetFilterKeyValue(cmbGoodsProfitRating.EditValue), true);

            //рейтинги продаж
            var goodsSaleRating = FilterUtility.GetSaleRatingEditValues(cmbGoodSaleRating.Text);
            var partSaleRating = FilterUtility.GetSaleRatingEditValues(cmbPartSaleRating.Text);

            OlapUtility.SetPivotFieldIdFilter(pivotGrid, "[Рейтинг продаж].[ABC по рейтингу товара].[ABC по рейтингу товара]", FilterUtility.GetFilterKeyValue(goodsSaleRating[true]), true);
            OlapUtility.SetPivotFieldIdFilter(pivotGrid, "[Рейтинг продаж региональный].[ABC по рейт. тов. рег].[ABC по рейт. тов. рег]", FilterUtility.GetFilterKeyValue(goodsSaleRating[false]), true);
            OlapUtility.SetPivotFieldIdFilter(pivotGrid, "[Рейтинг продаж].[ABC по рейтингу детали].[ABC по рейтингу детали]", FilterUtility.GetFilterKeyValue(partSaleRating[true]), true);
            OlapUtility.SetPivotFieldIdFilter(pivotGrid, "[Рейтинг продаж региональный].[ABC по рейт. дет. рег].[ABC по рейт. дет. рег]", FilterUtility.GetFilterKeyValue(partSaleRating[false]), true);

            OlapUtility.SetPivotFieldIdFilter(pivotGrid, "[XYZ товара].[Ключ].[Ключ]", FilterUtility.GetFilterKeyValue(cmbXyz.EditValue), true);

            OlapUtility.SetPivotFieldIdFilter(pivotGrid, "[Основание закупки].[Основание закупки].[Основание закупки]", FilterUtility.GetFilterKeyValue(cmbBasisOfPurchase.EditValue), true);

            OlapUtility.SetPivotFieldIdFilter(pivotGrid, "[Тип заказа].[Тип заказа ключ].[Тип заказа ключ]", FilterUtility.GetFilterKeyValue(cmbOrderType.EditValue), true);

            OlapUtility.SetPivotFieldIdFilter(pivotGrid, "[Номенклатура-Характеристика товара].[Ключ характеристика товара].[Ключ характеристика товара]",
                      FilterUtility.GetFilterKeyValue(cmbGoodsSpecifications.EditValue), true);

            OlapUtility.SetPivotFieldIdFilter(pivotGrid, "[Номенклатура-Характеристика детали].[Ключ характеристика детали].[Ключ характеристика детали]",
                      FilterUtility.GetFilterKeyValue(cmbPartSpecifications.EditValue), true);

            OlapUtility.SetPivotFieldIdFilter(pivotGrid, "[Операция продажи].[Ключ].[Ключ]", FilterUtility.GetFilterKeyValue(cmbOperation.EditValue), true);
            OlapUtility.SetPivotFieldIdFilter(pivotGrid, "[Номенклатура].[КлючНеассортиментный].[КлючНеассортиментный]", FilterUtility.GetFilterKeyValue(cmbAssortimentGoods.EditValue), true);

            DoApplyFilterToOlap(pivotGrid);
        }

        //установить значения фильтров в области фильтров в таблице на основании значений фильтров на панели поиска
        private void DoApplyFilterToOlap(PivotGridControl pvGrid)
        {
            // название - txtName - part_name - [Наименование].[Наименование].[Наименование] - LIKE part_name%
            var fieldPartName = pvGrid.Fields["[Номенклатура].[Название детали].[Название детали]"];

            OlapUtility.MoveFieldToFilterArea(fieldPartName);

            fieldPartName.FilterValues.FilterType = PivotFilterType.Excluded;
            fieldPartName.FilterValues.Clear();

            if (!string.IsNullOrEmpty(txtName.Text))
            {
                var partNamePattern = txtName.Text.ToLower();
                var valuesToInclude = new List<object>();

                foreach (var filterValue in fieldPartName.FilterValues.ValuesIncluded)
                {
                    var stringFilterValue = filterValue.ToString();
                    // Строгое совпадение (35687)
                    if (stringFilterValue.ToLower().Equals(partNamePattern))
                    {
                        valuesToInclude.Add(stringFilterValue);
                    }
                }

                fieldPartName.FilterValues.SetValues(valuesToInclude.ToArray(), PivotFilterType.Included, false);
            }


            var fieldSkubaNum = pvGrid.Fields["[Номенклатура].[SKN].[SKN]"];
            OlapUtility.MoveFieldToFilterArea(fieldSkubaNum);

            fieldSkubaNum.FilterValues.FilterType = PivotFilterType.Excluded;
            fieldSkubaNum.FilterValues.Clear();

            var sknValues = getMultiTextEditorValues(txtSKNNumber);
            if (sknValues.Any())
                fieldSkubaNum.FilterValues.SetValues(sknValues.ToArray(), PivotFilterType.Included, false);

            //№ уникальный детали
            var txtNumUnikue = pvGrid.Fields["[Номенклатура].[№ уникальный детали].[№ уникальный детали]"];

            OlapUtility.MoveFieldToFilterArea(txtNumUnikue);

            txtNumUnikue.FilterValues.FilterType = PivotFilterType.Excluded;
            txtNumUnikue.FilterValues.Clear();

            var uniqueNumberValues = getMultiTextEditorValues(txtUniqueNumber);
            if (uniqueNumberValues.Any())
                txtNumUnikue.FilterValues.SetValues(uniqueNumberValues.ToArray(), PivotFilterType.Included, false);

            //Код детали
            var txtPrtCode = pvGrid.Fields["[Номенклатура].[Код детали].[Код детали]"];

            OlapUtility.MoveFieldToFilterArea(txtPrtCode);

            txtPrtCode.FilterValues.FilterType = PivotFilterType.Excluded;
            txtPrtCode.FilterValues.Clear();

            var partCodeValues = getMultiTextEditorValues(txtPartCode);
            if (partCodeValues.Any())
                txtPrtCode.FilterValues.SetValues(partCodeValues.ToArray(), PivotFilterType.Included, false);

            //Код товара
            var txtGoodCode = pvGrid.Fields["[Номенклатура].[Код товара].[Код товара]"];

            OlapUtility.MoveFieldToFilterArea(txtGoodCode);

            txtGoodCode.FilterValues.FilterType = PivotFilterType.Excluded;
            txtGoodCode.FilterValues.Clear();

            var goodsCodeValues = getMultiTextEditorValues(txtGoodsCode);
            if (goodsCodeValues.Any())
                txtGoodCode.FilterValues.SetValues(goodsCodeValues.ToArray(), PivotFilterType.Included, false);
        }

        private IEnumerable<string> getMultiTextEditorValues(TextMultiSelectFromExcelButtonEdit multiTextEditor)
        {
            var multivalue = multiTextEditor.EditValue?.ToString()?.Trim() ?? "";
            return multivalue
                .Split(multiTextEditor.Separator)
                .Select(n => n.Trim());
        }

        //установить значения фильтров в области фильтров в таблице графика на основании значений фильтров на панели поиска
        private void DoApplyFilterToOlapChart()
        {
            // название - txtName - part_name - [Наименование].[Наименование].[Наименование] - LIKE part_name%
            var fieldPartName = pivotGridControlForChart.Fields["[Номенклатура].[Название детали].[Название детали]"];

            OlapUtility.MoveFieldToFilterArea(fieldPartName);

            fieldPartName.FilterValues.FilterType = PivotFilterType.Excluded;
            fieldPartName.FilterValues.Clear();

            if (!string.IsNullOrEmpty(txtName.Text))
            {
                var partNamePattern = txtName.Text.ToLower();
                var valuesToInclude = new List<object>();

                foreach (var filterValue in fieldPartName.FilterValues.ValuesIncluded)
                {
                    var stringFilterValue = filterValue.ToString();
                    // Строгое совпадение (35687)
                    if (stringFilterValue.ToLower().Equals(partNamePattern))
                    {
                        valuesToInclude.Add(stringFilterValue);
                    }
                }

                fieldPartName.FilterValues.SetValues(valuesToInclude.ToArray(), PivotFilterType.Included, false);
            }


            var fieldSkubaNum = pivotGridControlForChart.Fields["[Номенклатура].[SKN].[SKN]"];
            OlapUtility.MoveFieldToFilterArea(fieldSkubaNum);

            fieldSkubaNum.FilterValues.FilterType = PivotFilterType.Excluded;
            fieldSkubaNum.FilterValues.Clear();

            var sknValues = getMultiTextEditorValues(txtSKNNumber);
            if (sknValues.Any())
                fieldSkubaNum.FilterValues.SetValues(sknValues.ToArray(), PivotFilterType.Included, false);

            //№ уникальный детали
            var txtNumUnikue = pivotGridControlForChart.Fields["[Номенклатура].[№ уникальный детали].[№ уникальный детали]"];
            OlapUtility.MoveFieldToFilterArea(txtNumUnikue);

            txtNumUnikue.FilterValues.FilterType = PivotFilterType.Excluded;
            txtNumUnikue.FilterValues.Clear();

            var uniqueNumberValues = getMultiTextEditorValues(txtUniqueNumber);
            if (uniqueNumberValues.Any())
                txtNumUnikue.FilterValues.SetValues(uniqueNumberValues.ToArray(), PivotFilterType.Included, false);

            //Код детали
            var txtPrtCode = pivotGridControlForChart.Fields["[Номенклатура].[Код детали].[Код детали]"];
            OlapUtility.MoveFieldToFilterArea(txtPrtCode);

            txtPrtCode.FilterValues.FilterType = PivotFilterType.Excluded;
            txtPrtCode.FilterValues.Clear();

            var partCodeValues = getMultiTextEditorValues(txtPartCode);
            if (partCodeValues.Any())
                txtPrtCode.FilterValues.SetValues(partCodeValues.ToArray(), PivotFilterType.Included, false);

            //Код товара
            var txtGoodCode = pivotGridControlForChart.Fields["[Номенклатура].[Код товара].[Код товара]"];
            OlapUtility.MoveFieldToFilterArea(txtGoodCode);

            txtGoodCode.FilterValues.FilterType = PivotFilterType.Excluded;
            txtGoodCode.FilterValues.Clear();

            var goodsCodeValues = getMultiTextEditorValues(txtGoodsCode);
            if (goodsCodeValues.Any())
                txtGoodCode.FilterValues.SetValues(goodsCodeValues.ToArray(), PivotFilterType.Included, false);
        }

        public override List<FilterHistoryColumn> GetFilterHistoryColumns()
        {
            return new List<FilterHistoryColumn>
            {
                new FilterHistoryColumn(lblPeriodFrom.Text, txtPeriodFrom, 90),
                new FilterHistoryColumn(lblPeriodTo.Text, txtPeriodTo, 90),
                new FilterHistoryColumn(lblSKNNumber.Text, txtSKNNumber.Edit, 150),
                new FilterHistoryColumn(lblName.Text, txtName, 90),
                new FilterHistoryColumn(lblUniqueNumber.Text, txtUniqueNumber.Edit, 90),
                new FilterHistoryColumn(labelControl1.Text, txtPartCode.Edit, 150),
                new FilterHistoryColumn(labelControl2.Text, txtGoodsCode.Edit, 90),
                new FilterHistoryColumn(lblAssortimentGoods.Text, cmbAssortimentGoods, 75),
                new FilterHistoryColumn(labelControl3.Text, txtContract, 150),
                new FilterHistoryColumn(lblGroup.Text, cmbGroup, 150),
                new FilterHistoryColumn(lblOperation.Text, cmbOperation, 150),
                new FilterHistoryColumn(lblSubgroup.Text, cmbSubgroup, 150),
                new FilterHistoryColumn(lblProducer.Text, cmbProducer, 150),
                new FilterHistoryColumn(lblGoodSaleRating.Text, cmbGoodSaleRating, 150),
                new FilterHistoryColumn(lblPartSaleRating.Text, cmbPartSaleRating, 150),
                new FilterHistoryColumn(lblGoodsProfitRating.Text, cmbGoodsProfitRating, 150),
                new FilterHistoryColumn(lblBasisOfPurchase.Text, cmbBasisOfPurchase, 150),
                new FilterHistoryColumn(lblOrderType.Text, cmbOrderType, 150),
                new FilterHistoryColumn(lblAcceptanceStorage.Text, txtAccepStorage, 500),
                new FilterHistoryColumn(lblResponsibleManager.Text, txtResponsibleManager, 500),
                new FilterHistoryColumn(lblPaymentType.Text, _paymentTypeFilter.Control, 150)
            };
        }

        private void TrimTimeFromDateFields()
        {
            TrimTimeFromControl(txtPeriodFrom);
            TrimTimeFromControl(txtPeriodTo);
        }

        private static void TrimTimeFromControl(DateEdit control)
        {
            if (control.EditValue == null)
                return;

            if (control.DateTime.Hour == 0 && control.DateTime.Minute == 0 && control.DateTime.Second == 0 && control.DateTime.Second == 0 && control.DateTime.Millisecond == 0)
                return;

            control.EditValue = control.DateTime.Date;

        }

        //настройки смены профиля таблицы на график
        private void ApplyPresentation(bool iNeedRefresh)
        {
            var itemPresentation = (BarEditItem)_buttons[(int)ToolBarButtonIndex.Presentation].Item;
            var presentation = itemPresentation.EditValue.ToString();
            _presentation = presentation;

            var isChart = presentation == "График";
            using (new WaitCursor())
            {
                if (!isChart)
                {
                    // детачим кнопки грида для грида графика
                    if (_attachedButtons.Contains(ToolBarButtonIndex.GridLayoutSettings))
                    {
                        _attachedButtons.Remove(ToolBarButtonIndex.GridLayoutSettings);

                        DetachToolBarButton(_menuButtons, _buttons, ToolBarButtonIndex.GridLayoutSettings, null);

                        var item = (BarEditItem)_buttons[(int)ToolBarButtonIndex.GridLayoutSettings].Item;
                        item.SuperTip.Items.Clear();
                        item.EditValueChanged -= gridLayout_EditValueChanged;
                        GridLayoutSettingsUtility.DetachBarEditItem(item);
                    }

                    // аттачим кнопки грида
                    if (!_attachedButtons.Contains(ToolBarButtonIndex.GridLayoutSettings))
                    {
                        AttachToolBarButton(_menuButtons, _buttons, ToolBarButtonIndex.GridLayoutSettings, null);

                        var item = (BarEditItem)_buttons[(int)ToolBarButtonIndex.GridLayoutSettings].Item;
                        var titleItem = new ToolTipTitleItem();
                        titleItem.Text = "Настройки таблицы";
                        if (item.SuperTip == null)
                            item.SuperTip = new SuperToolTip();
                        item.SuperTip.Items.Add(titleItem);

                        GridLayoutSettingsUtility.AttachBarEditItem(item, GridLayoutSettingsCode, null, null, pivotGrid, _lastGridSettings);
                        // если открываем форму, то сохраняем состояние, иначе если уже есть сохранённое, то загружаем
                        if (_lastGridSettings != null)
                        {
                            if (item.EditValue != _lastGridSettings)
                            {
                                item.EditValue = _lastGridSettings;
                                if (iNeedRefresh)
                                    GridLayoutSettingsUtility.RefreshPivotGrid(GridLayoutSettingsCode, pivotGrid);
                            }
                        }
                        else
                        {
                            _lastGridSettings = item.EditValue;
                            if (iNeedRefresh)
                                GridLayoutSettingsUtility.RefreshPivotGrid(GridLayoutSettingsCode, pivotGrid);
                        }

                        item.EditValueChanged += gridLayout_EditValueChanged;
                        _attachedButtons.Add(ToolBarButtonIndex.GridLayoutSettings);
                    }

                    SelectTabPage(tbpGrid);
                }
                else if (isChart)
                {

                    // детачим кнопки грида
                    if (_attachedButtons.Contains(ToolBarButtonIndex.GridLayoutSettings))
                    {
                        _attachedButtons.Remove(ToolBarButtonIndex.GridLayoutSettings);

                        DetachToolBarButton(_menuButtons, _buttons, ToolBarButtonIndex.GridLayoutSettings, null);

                        var item = (BarEditItem)_buttons[(int)ToolBarButtonIndex.GridLayoutSettings].Item;
                        item.SuperTip.Items.Clear();
                        item.EditValueChanged -= gridLayout_EditValueChanged;
                        GridLayoutSettingsUtility.DetachBarEditItem(item);
                    }
                    // аттачим кнопки грида для графика
                    if (!_attachedButtons.Contains(ToolBarButtonIndex.GridLayoutSettings))
                    {
                        AttachToolBarButton(_menuButtons, _buttons, ToolBarButtonIndex.GridLayoutSettings, null);

                        var item = (BarEditItem)_buttons[(int)ToolBarButtonIndex.GridLayoutSettings].Item;
                        var titleItem = new ToolTipTitleItem();
                        titleItem.Text = "Настройки таблицы";
                        if (item.SuperTip == null)
                            item.SuperTip = new SuperToolTip();
                        item.SuperTip.Items.Add(titleItem);

                        GridLayoutSettingsUtility.AttachBarEditItem(item, GridChartLayoutSettingsCode, null, null, pivotGridControlForChart, _lastGridChartSettings);
                        // если открываем форму, то сохраняем состояние, иначе если уже есть сохранённое, то загружаем
                        if (_lastGridChartSettings != null)
                        {
                            if (item.EditValue != _lastGridChartSettings)
                            {
                                item.EditValue = _lastGridChartSettings;
                                if (iNeedRefresh)
                                {
                                    GridLayoutSettingsUtility.RefreshPivotGrid(GridChartLayoutSettingsCode, pivotGridControlForChart);
                                }
                            }
                        }
                        else
                        {
                            _lastGridSettings = item.EditValue;
                            if (iNeedRefresh)
                                GridLayoutSettingsUtility.RefreshPivotGrid(GridChartLayoutSettingsCode, pivotGridControlForChart);
                        }

                        item.EditValueChanged += gridLayout_EditValueChanged;
                        _attachedButtons.Add(ToolBarButtonIndex.GridLayoutSettings);

                    }
                    SelectTabPage(tbpChart);
                }

                RegistryUtility.RemainsPresentation = presentation;
            }

            RegistryUtility.RemainsPresentation = presentation;
        }


        private void SelectTabPage(XtraTabPage tabPage)
        {
            if (tabChilds.SelectedTabPage == tabPage)
                return;

            tabChilds.SelectedTabPage = tabPage;
        }

        #region Chart

        private void InitChart()
        {
            chartReport.Legend.AlignmentHorizontal = LegendAlignmentHorizontal.Left;
            chartReport.Legend.BackColor = new Color();

            chartReport.ObjectSelected += chartReport_ObjectSelected;
            chartReport.ObjectHotTracked += chartReport_ObjectHotTracked;
            ((XYDiagram)chartReport.Diagram).AxisY.Label.TextPattern = "{V:N}";

            ChartUtility.ConfigureChartSelection(chartReport);

            chartReport.PivotGridDataSourceOptions.AutoLayoutSettingsEnabled = false;
            ((XYDiagram)chartReport.Diagram).AxisX.DateTimeScaleOptions.MeasureUnit = DateTimeMeasureUnit.Day;
            ((XYDiagram)chartReport.Diagram).AxisX.DateTimeScaleOptions.GridAlignment = DateTimeGridAlignment.Day;

            chartControlOlap.Legend.AlignmentHorizontal = LegendAlignmentHorizontal.Left;
            chartControlOlap.Legend.BackColor = new Color();
            ChartUtility.ConfigureChartSelection(chartControlOlap);
            chartControlOlap.VisibleChanged += chart_VisibleChanged;
            ((XYDiagram)chartControlOlap.Diagram).AxisY.Label.TextPattern = "{V:N}";
            _isChartInitialized = false;
        }

        private void BindChart()
        {
            TrimTimeFromDateFields();
            if (IsAnyFilterParamSet)
                TraceSearch();

            OlapUtility.SetPivotFieldDateFilter(pivotGridControlForChart, DateFilter, txtPeriodFrom.EditValue, txtPeriodTo.EditValue);

            OlapUtility.SetPivotFieldIdFilter(pivotGridControlForChart, "[Номенклатура-Группа].[Ключ].[Ключ]", FilterUtility.GetFilterKeyValue(cmbGroup.EditValue), true);

            // Подгруппа  cmbSubgroup - id_subgroup - 
            var subgroupFilterValue = FilterUtility.GetSubgroupFilter(cmbSubgroup.EditValue);
            OlapUtility.SetPivotFieldIdFilter(pivotGridControlForChart, "[Номенклатура-Подгруппа].[Код подгруппы].[Код подгруппы]", subgroupFilterValue, true);
            if (_isChartInitialized)
            {
                // Склад   txtStorage - storages_list - 
                var ShpStorageFilterValue = FilterUtility.GetStorageFilter(txtAccepStorage.EditValue);
                OlapUtility.SetPivotFieldIdFilter(pivotGridControlForChart, "[Склад].[Ключ абстрактный].[Ключ абстрактный]", ShpStorageFilterValue, true);

                // Тип операции
                // cmbPaymentType - payment_type_list - [Тип операции].[Ключ].[Ключ]
                var paymentTypeFilter = FilterUtility.GetFilterKeyValue(_paymentTypeFilter.Control.EditValue);
                if (paymentTypeFilter == null)
                    paymentTypeFilter = new object[0];

                OlapUtility.SetPivotFieldIdFilter(pivotGridControlForChart, "[Тип операции].[Ключ].[Ключ]", paymentTypeFilter, true);
            }

            OlapUtility.SetPivotFieldIdFilter(pivotGridControlForChart, "[Поставщик].[Ключ].[Ключ]", FilterUtility.GetFilterKeyValue(txtContract.EditValue), true);

            // cmbProducer 
            var producerFilterValue = FilterUtility.GetProducerFilter(cmbProducer.EditValue);
            OlapUtility.SetPivotFieldIdFilter(pivotGridControlForChart, "[Номенклатура].[ПроизводительКлюч].[ПроизводительКлюч]", producerFilterValue, true);

            OlapUtility.SetPivotFieldIdFilter(pivotGridControlForChart, "[Ответственный за подгруппу].[Id_abstract_contractor].[Id_abstract_contractor]", FilterUtility.GetFilterKeyValue(txtResponsibleManager.EditValue), true);
            OlapUtility.SetPivotFieldIdFilter(pivotGridControlForChart, "[Ответственный за филиал].[id_responsible_abstract].[id_responsible_abstract]", FilterUtility.GetFilterKeyValue(txtResponsibleForDeliveryToBranch.EditValue), true);
            
            OlapUtility.SetPivotFieldIdFilter(pivotGridControlForChart, "[Рейтинг по прибыли].[ABC по рейтингу товара по прибыли].[ABC по рейтингу товара по прибыли]", FilterUtility.GetFilterKeyValue(cmbGoodsProfitRating.EditValue), true);

            //рейтинги продаж
            var goodsSaleRating = FilterUtility.GetSaleRatingEditValues(cmbGoodSaleRating.Text);
            var partSaleRating = FilterUtility.GetSaleRatingEditValues(cmbPartSaleRating.Text);

            OlapUtility.SetPivotFieldIdFilter(pivotGridControlForChart, "[Рейтинг продаж].[ABC по рейтингу товара].[ABC по рейтингу товара]", FilterUtility.GetFilterKeyValue(goodsSaleRating[true]), true);
            OlapUtility.SetPivotFieldIdFilter(pivotGridControlForChart, "[Рейтинг продаж региональный].[ABC по рейт. тов. рег].[ABC по рейт. тов. рег]", FilterUtility.GetFilterKeyValue(goodsSaleRating[false]), true);
            OlapUtility.SetPivotFieldIdFilter(pivotGridControlForChart, "[Рейтинг продаж].[ABC по рейтингу детали].[ABC по рейтингу детали]", FilterUtility.GetFilterKeyValue(partSaleRating[true]), true);
            OlapUtility.SetPivotFieldIdFilter(pivotGridControlForChart, "[Рейтинг продаж региональный].[ABC по рейт. дет. рег].[ABC по рейт. дет. рег]", FilterUtility.GetFilterKeyValue(partSaleRating[false]), true);

            OlapUtility.SetPivotFieldIdFilter(pivotGridControlForChart, "[XYZ товара].[Ключ].[Ключ]", FilterUtility.GetFilterKeyValue(cmbXyz.EditValue), true);

            OlapUtility.SetPivotFieldIdFilter(pivotGridControlForChart, "[Номенклатура-Характеристика товара].[Ключ характеристика товара].[Ключ характеристика товара]",
                      FilterUtility.GetFilterKeyValue(cmbGoodsSpecifications.EditValue), true);

            OlapUtility.SetPivotFieldIdFilter(pivotGridControlForChart, "[Номенклатура-Характеристика детали].[Ключ характеристика детали].[Ключ характеристика детали]",
                      FilterUtility.GetFilterKeyValue(cmbPartSpecifications.EditValue), true);

            OlapUtility.SetPivotFieldIdFilter(pivotGridControlForChart, "[Основание закупки].[Основание закупки].[Основание закупки]", FilterUtility.GetFilterKeyValue(cmbBasisOfPurchase.EditValue), true);
            OlapUtility.SetPivotFieldIdFilter(pivotGridControlForChart, "[Тип заказа].[Тип заказа ключ].[Тип заказа ключ]", FilterUtility.GetFilterKeyValue(cmbOrderType.EditValue), true);
            OlapUtility.SetPivotFieldIdFilter(pivotGridControlForChart, "[Операция продажи].[Ключ].[Ключ]", FilterUtility.GetFilterKeyValue(cmbOperation.EditValue), true);
            OlapUtility.SetPivotFieldIdFilter(pivotGridControlForChart, "[Номенклатура].[КлючНеассортиментный].[КлючНеассортиментный]", FilterUtility.GetFilterKeyValue(cmbAssortimentGoods.EditValue), true);

            DoApplyFilterToOlapChart();
            _isChartInitialized = true;
        }

        private void RefreshAutoChart()
        {
            ChartUtility.RefreshAutoChart(pivotGrid, chartControlOlap, _isPointsLabelVisible);
        }
        #endregion

        #region IToolBarForm Members

        public override void AttachToolBarButtons(BarItemLinkCollection menuButtons, BarItemLinkCollection buttons, out bool isAnyButtonVisible)
        {
            base.AttachToolBarButtons(menuButtons, buttons, out isAnyButtonVisible);

            InitForm();

            ((BarCheckItem)menuButtons[(int)ToolBarButtonIndex.LogReport].Item).Checked = false;
            ((BarCheckItem)buttons[(int)ToolBarButtonIndex.LogReport].Item).Checked = false;

            _menuButtons = menuButtons;
            _buttons = buttons;

            AttachToolBarButton(menuButtons, buttons, ToolBarButtonIndex.UseFilter, ToolBarButtonUseFilter_ItemClick);
            AttachToolBarButton(menuButtons, buttons, ToolBarButtonIndex.ResetFilter, ToolBarButtonResetFilter_ItemClick);
            AttachToolBarButton(menuButtons, buttons, ToolBarButtonIndex.ShowFilterHistory, ToolBarButtonShowFilterHistory_ItemClick);
            AttachToolBarButton(menuButtons, buttons, ToolBarButtonIndex.ToggleFilter, ToolBarButtonToggleFilter_ItemClick);

            AttachToolBarButton(menuButtons, buttons, ToolBarButtonIndex.ChartLegent, ToolBarButtonChartLegent_ItemClick);
            AttachToolBarButton(menuButtons, buttons, ToolBarButtonIndex.ChartPointLabel, ToolBarButtonChartPointLabel_ItemClick);
            AttachToolBarButton(menuButtons, buttons, ToolBarButtonIndex.Presentation, null);

            if (FormModuleObject.IsDownloadSupported && ProfileUtility.CurrentUser.CanDownloadEntity(FormModuleObject))
                AttachToolBarButton(menuButtons, buttons, ToolBarButtonIndex.Download, ToolBarButtonDownload_ItemClick);

            AttachToolBarButton(menuButtons, buttons, ToolBarButtonIndex.Info, null);
            var infoMenu = (BarStaticItem)menuButtons[(int)ToolBarButtonIndex.Info].Item;
            infoMenu.Visibility = BarItemVisibility.Never;

            var itemPresentation = (BarEditItem)buttons[(int)ToolBarButtonIndex.Presentation].Item;

            if (itemPresentation.SuperTip == null)
            {
                var titlePresentation = new ToolTipTitleItem();
                titlePresentation.Text = "Представление";
                itemPresentation.SuperTip = new SuperToolTip();
                itemPresentation.SuperTip.Items.Add(titlePresentation);
            }


            var presentationLookupEdit = (RepositoryItemLookUpEdit)itemPresentation.Edit;
            ListControlUtility.ConfigureLookUpEditEvents(presentationLookupEdit);
            presentationLookupEdit.NullText = string.Empty;

            var dtPresentation = new DataTable();
            dtPresentation.Columns.Add("Value");

            presentationLookupEdit.ShowHeader = false;
            presentationLookupEdit.DataSource = dtPresentation;

            dtPresentation.Rows.Add("Сводная таблица");
            dtPresentation.Rows.Add("График");

            var defaultValue = _presentation ?? (string.IsNullOrEmpty(RegistryUtility.RemainsPresentation) ? "Сводная таблица" : RegistryUtility.RemainsPresentation);

            itemPresentation.EditValue = _presentation;
            if (itemPresentation.EditValue.IsNullOrDBNullOrEmpty())
            {
                itemPresentation.EditValue = defaultValue;
            }

            itemPresentation.EditValueChanged += itemPresentation_EditValueChanged;

            presentationLookupEdit.ValueMember = "Value";
            presentationLookupEdit.DisplayMember = "Value";

            ApplyPresentation(false);
        }


        public override void DetachToolBarButtons(BarItemLinkCollection menuButtons, BarItemLinkCollection buttons)
        {
            base.DetachToolBarButtons(menuButtons, buttons);

            if (_attachedButtons.Contains(ToolBarButtonIndex.GridLayoutSettings))
            {
                _attachedButtons.Remove(ToolBarButtonIndex.GridLayoutSettings);

                DetachToolBarButton(menuButtons, buttons, ToolBarButtonIndex.GridLayoutSettings, null);

                var item = (BarEditItem)buttons[(int)ToolBarButtonIndex.GridLayoutSettings].Item;
                item.SuperTip.Items.Clear();
                item.EditValueChanged -= gridLayout_EditValueChanged;
                GridLayoutSettingsUtility.DetachBarEditItem(item);
            }

            if (_attachedButtons.Contains(ToolBarButtonIndex.ChartReportDateType))
            {
                _attachedButtons.Remove(ToolBarButtonIndex.ChartReportDateType);

                DetachToolBarButton(_menuButtons, _buttons, ToolBarButtonIndex.ChartReportDateType, null);
            }

            DetachToolBarButton(menuButtons, buttons, ToolBarButtonIndex.UseFilter, ToolBarButtonUseFilter_ItemClick);
            DetachToolBarButton(menuButtons, buttons, ToolBarButtonIndex.ResetFilter, ToolBarButtonResetFilter_ItemClick);
            DetachToolBarButton(menuButtons, buttons, ToolBarButtonIndex.ShowFilterHistory, ToolBarButtonShowFilterHistory_ItemClick);
            DetachToolBarButton(menuButtons, buttons, ToolBarButtonIndex.ToggleFilter, ToolBarButtonToggleFilter_ItemClick);
            DetachToolBarButton(menuButtons, buttons, ToolBarButtonIndex.ChartLegent, ToolBarButtonChartLegent_ItemClick);
            DetachToolBarButton(menuButtons, buttons, ToolBarButtonIndex.ChartPointLabel, ToolBarButtonChartPointLabel_ItemClick);
            DetachToolBarButton(menuButtons, buttons, ToolBarButtonIndex.Presentation, null);
            DetachToolBarButton(menuButtons, buttons, ToolBarButtonIndex.Download, ToolBarButtonDownload_ItemClick);
            DetachToolBarButton(menuButtons, buttons, ToolBarButtonIndex.Info, null);

            var itemPresentation = (BarEditItem)buttons[(int)ToolBarButtonIndex.Presentation].Item;
            itemPresentation.EditValueChanged -= itemPresentation_EditValueChanged;
        }

        #endregion

        #region Event handlers

        // загрузка формы - установка фильтров периода, если они пустые
        protected override void OnFormLoad()
        {
            base.OnFormLoad();

            ConfigureAndPopulateLookUps();
            InitForm();
            foreach (Control c in pnlFilter.Controls)
            {
                if (c is BaseEdit edit)
                    edit.KeyUp += RemainsTabularTelerikReportFormKeyUp;
            }

            var dt = GridFormSearchTrace.GetGridFormSearchTraceDataByModuleObject(FormModuleObject.Id.Value);
            if (dt != null && dt.Rows.Count > 0)
                FilterUtility.SetFiltersByHistory((int)dt.Rows[0]["id"], pnlFilter);

            if (txtPeriodTo.EditValue == null)
                txtPeriodTo.EditValue = DateTime.Today.AddDays(-1).Date;

            if (txtPeriodFrom.EditValue == null)
                txtPeriodFrom.EditValue = DateTime.Today.AddMonths(-1).Date;

            if (tabChilds.SelectedTabPage == tbpChart)
                OlapUtility.SetPivotFieldDateFilter(pivotGridControlForChart, DateFilter, txtPeriodFrom.EditValue, txtPeriodTo.EditValue);
            else if (tabChilds.SelectedTabPage == tbpGrid)
                OlapUtility.SetPivotFieldDateFilter(pivotGrid, DateFilter, txtPeriodFrom.EditValue, txtPeriodTo.EditValue);

            ApplyFilter();
        }

        private void PopulateAssortmentGoodsCheckedCombo(CheckedComboBoxEdit cmb)
        {
            var dtAssortimentGoods = GoodsBlockingReason.SelectTable();
            var dr = dtAssortimentGoods.NewRow();
            dr[GoodsBlockingReason.BaseColumns.Id] = 0;
            dr[GoodsBlockingReason.Columns.Name] = "Ассортиментный товар";
            dtAssortimentGoods.Rows.Add(dr);
            cmb.Properties.DataSource = new DataView(dtAssortimentGoods, null, string.Format("{0} ASC", GoodsBlockingReason.BaseColumns.Id), DataViewRowState.CurrentRows);
            cmb.Properties.DisplayMember = GoodsBlockingReason.Columns.Name;
            cmb.Properties.ValueMember = GoodsBlockingReason.BaseColumns.Id;
            cmb.Properties.NullText = string.Empty;
            cmb.Properties.SelectAllItemVisible = true;
            cmb.Properties.PopupFormMinSize = new Size(120, cmb.Properties.PopupFormMinSize.Height);
        }

        //инициализация всех свойств и обработчиков событий формы на закладках
        private void InitForm()
        {
            if (_isFormInitialized)
                return;

            _isFormInitialized = true;

            InitialiseCube();
            InitGrid();
            InitChart();
        }

        // обработка кнопки "скинуть значения всех фильтров" - обнуляет все текстовые и ставит период прошлый месяц до вчерашнего дня
        private void ToolBarButtonResetFilter_ItemClick(object sender, ItemClickEventArgs e)
        {
            ResetFilter();
        }

        // обработка кнопки "применит фильтр" - применение фильтров и обновление pivotTable
        private void ToolBarButtonUseFilter_ItemClick(object sender, ItemClickEventArgs e)
        {
            using (new WaitCursor())
            {
                ApplyFilter();
            }
        }

        private void ToolBarButtonShowFilterHistory_ItemClick(object sender, ItemClickEventArgs e)
        {
            var filterColumns = GetFilterHistoryColumns();
            var traceId = -1;
            using (var form = new FilterHistoryModalForm(null, FormModuleObject.Id, filterColumns))
            {
                FormUtility.ShowDialog(form);
                if (form.DialogResult == DialogResult.Cancel)
                    return;

                traceId = form.SelectedGridFormSearchTraceId;
            }
            if (traceId > -1)
                FilterUtility.SetFiltersByHistory(traceId, pnlFilter);

            using (new WaitCursor())
                ApplyFilter();
        }

        private void ToolBarButtonToggleFilter_ItemClick(object sender, ItemClickEventArgs e)
        {
            pnlFilter.Visible = !pnlFilter.Visible;
        }

        private void itemPresentation_EditValueChanged(object sender, EventArgs e)
        {
            ApplyPresentation(false);
        }

        private void pivotGrid_FieldValueExpanding(object sender, PivotFieldValueCancelEventArgs e)
        {
            StartWaitCursor();
        }

        private void pivotGrid_FieldValueNotExpanded(object sender, PivotFieldValueEventArgs e)
        {
            EndWaitCursor();
        }

        private void pivotGrid_FieldValueCollapsed(object sender, PivotFieldValueEventArgs e)
        {
            pivotGrid.Cells.Selection = Rectangle.Empty;
            RefreshAutoChart();
        }

        private void pivotGridControlForChart_FieldValueCollapsed(object sender, PivotFieldValueEventArgs e)
        {
            RefreshAutoChartReport();
        }

        private void pivotGrid_FieldValueExpanded(object sender, PivotFieldValueEventArgs e)
        {
            pivotGrid.Cells.Selection = Rectangle.Empty;
            RefreshAutoChart();
            EndWaitCursor();
        }

        private void pivotGridControlForChart_FieldValueExpanded(object sender, PivotFieldValueEventArgs e)
        {
            RefreshAutoChartReport();
            EndWaitCursor();
        }

        private void pivotGrid_CellClick(object sender, PivotCellEventArgs e)
        {
            RefreshAutoChart();
        }

        private void pivotGridControlForChart_CellClick(object sender, PivotCellEventArgs e)
        {
            RefreshAutoChartReport();
        }

        private void pivotGrid_CellSelectionChanged(object sender, EventArgs e)
        {
            RefreshAutoChart();
        }

        private void pivotGridControlForChart_CellSelectionChanged(object sender, EventArgs e)
        {
            RefreshAutoChartReport();
        }

        private void gridLayout_EditValueChanged(object sender, EventArgs e)
        {
            var olapPivot = (BarEditItem)_buttons[(int)ToolBarButtonIndex.OlapPivotGridLayoutSettings].Item;
            _lastGridSettings = olapPivot.EditValue;
        }

        // экспорт в Excel стандартная форма
        private void ToolBarButtonDownload_ItemClick(object sender, ItemClickEventArgs e)
        {
            using (var dlg = new SaveFileDialog())
            {
                var fileName = string.Format("Комплексный_отчёт_по_товародвижению_{0:d}_{1}", DateTime.Now,
                    User.GetUserBriefName(ProfileUtility.CurrentUser.Id.Value));

                var xls = new XlsxExportOptionsEx();

                if (tabChilds.SelectedTabPage == tbpGrid)
                {
                    xls.SheetName = "Комплексный отчёт по товародвижению (сводный)";
                    fileName += "_(сводный).xlsx";
                    ExportUtility.ExportToXlsx(pivotGrid, fileName, xls, ControlFiltersNamesForSheetHeaders, GetSearchTrace());
                }
            }
        }

        private void pivotGrid_Click(object sender, EventArgs e)
        {
            RefreshAutoChart();
        }

        //обработка смены закладок. пока их две- таблица и график
        private void tabChilds_SelectedPageChanged(object sender, TabPageChangedEventArgs e)
        {
            HandleSelectedTabChanged();
        }

        // при смене закладок прячем кастомную форму старого pivottable и делаем привязку к установленным фильтрам нового pivottable
        private void HandleSelectedTabChanged()
        {
            if (tabChilds.SelectedTabPage == tbpGrid)
            {
                BindData();
                if (pivotGridControlForChart.CustomizationForm == null)
                    return;

                if (pivotGridControlForChart.CustomizationForm.Visible)
                    pivotGridControlForChart.CustomizationForm.Close();
            }
            else if (tabChilds.SelectedTabPage == tbpChart)
            {
                BindChart();
                if (pivotGrid.CustomizationForm == null)
                    return;

                if (pivotGrid.CustomizationForm.Visible)
                    pivotGrid.CustomizationForm.Close();
            }
        }

        private void RefreshAutoChartReport()
        {
            ChartUtility.ChangeChartPointLabelVisibility(chartReport, _isPointsLabelVisible);
        }

        private void pivotGridControlForChart_Click(object sender, EventArgs e)
        {
            RefreshAutoChartReport();
        }

        private void ToolBarButtonChartLegent_ItemClick(object sender, ItemClickEventArgs e)
        {
            _isLegendVisible = !_isLegendVisible;

            ChartUtility.ChangeChartLegendVisibility(chartControlOlap, _isLegendVisible);
            ChartUtility.ChangeChartLegendVisibility(chartReport, _isLegendVisible);

            RegistryUtility.IsRemainsReportChartLegendVisible = _isLegendVisible;
        }

        private void ToolBarButtonChartPointLabel_ItemClick(object sender, ItemClickEventArgs e)
        {
            _isPointsLabelVisible = !_isPointsLabelVisible;

            ChartUtility.ChangeChartPointLabelVisibility(chartControlOlap, _isPointsLabelVisible);
            ChartUtility.ChangeChartPointLabelVisibility(chartReport, _isPointsLabelVisible);

            RegistryUtility.IsRemainsReportChartPointLabelVisible = _isPointsLabelVisible;
        }

        private void chart_VisibleChanged(object sender, EventArgs e)
        {
            var chart = (ChartControl)sender;
            if (chart.Visible)
            {
                ChartUtility.ChangeChartLegendVisibility(chart, _isLegendVisible);
                ChartUtility.ChangeChartPointLabelVisibility(chart, _isPointsLabelVisible);
            }
        }

        private void chartReport_ObjectHotTracked(object sender, HotTrackEventArgs e)
        {
            if (!(e.Object is Series) && !(e.Object is PointSeriesLabel) && !(e.Object is SideBySideBarSeriesLabel))
            {
                e.Cancel = true;
            }
        }

        private void chartReport_ObjectSelected(object sender, HotTrackEventArgs e)
        {
            if (!(e.Object is Series) && !(e.Object is PointSeriesLabel) && !(e.Object is SideBySideBarSeriesLabel))
            {
                e.Cancel = true;
                chartReport.ClearSelection(false);
            }
        }

        private void RemainsTabularTelerikReportFormKeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode != Keys.Enter)
                return;

            using (new WaitCursor())
            {
                if (tabChilds.SelectedTabPage == tbpChart)
                    BindChart();
                else if (tabChilds.SelectedTabPage == tbpGrid)
                    BindData();
            }
        }

        private void cmbSubgroup_CustomDisplayText(object sender, DevExpress.XtraEditors.Controls.CustomDisplayTextEventArgs e)
        {
            ListControlUtility.QueryCheckedComboboxDisplayValue(e);
        }

        private void RemainsTabularTelerikReportForm_Load(object sender, EventArgs e)
        {
            InitialiseCube();
        }

        private void InitialiseCube()
        {
            if (!pivotGrid.OLAPConnectionString.IsNullOrEmpty())
                return;

            // TODO: перенести в настройки ADTS
            var connectionString = "Data Source=srvt-sqlolap-01;Catalog=Remains_t;Cube=Остатки";
            // TODO: OlapUtility.IsCubeProcessed - не работает для обычных пользователей. Databases не вычитываются, пустой массив
            //if (!OlapUtility.IsCubeProcessed(connectionString, "Remains_t", "Остатки"))
            //{
            //    DialogUtility.ShowErrorMessage(this, "Сервис аналитики по товародвижению недоступен.");
            //    BeginInvoke(new MethodInvoker(Dispose));
            //}

            pivotGrid.OLAPConnectionString = pivotGridControlForChart.OLAPConnectionString = connectionString;

            // Автозагрузка полей (из перспективы, если указана)
            pivotGrid.RetrieveFields(PivotArea.DataArea, false);
            pivotGridControlForChart.RetrieveFields(PivotArea.DataArea, false);

            OlapUtility.SetPivotGridOlapMode(pivotGrid, FieldsForCustomSort, CompositeFieldsForCustomSort);
        }

        private void pivotGrid_AsyncOperationCompleted(object sender, EventArgs e)
        {
            if (_settings.GetValue(ApplicationSettingsRepository.IsReportOlapLogKey) == ApplicationSettingsRepository.IsReportOlapLogEnabledValue)
            {
                _isAsyncInProgress = pivotGrid.IsAsyncInProgress;

                if (!_isAsyncInProgress)
                {
                    loadManager.StopTimer();

                    OlapUtility.SetLogReport(Database, Text, _windowGuid, _sessionGuid);
                    _sessionGuid = OlapUtility.GetNextSessionId();
                }
            }
        }

        private void pivotGrid_OLAPQueryData(object sender, PivotOlapQueryDataEventArgs e)
        {
            var context = new DBContext(Database, true);

            if (_settings.GetValue(ApplicationSettingsRepository.IsReportOlapLogKey) == ApplicationSettingsRepository.IsReportOlapLogEnabledValue)
            {
                MainFormAccessor.Instance.Invoke(new MethodInvoker(() =>
                {
                    var log = new OlapReportLog(pnlFilter)
                    {
                        Presentation = _presentation,
                        GridSettings = ((BarEditItem)_buttons[(int)ToolBarButtonIndex.GridLayoutSettings].Item).EditValue,
                        PeriodFrom = ConversionUtility.GetNullableDateTime(txtPeriodFrom.EditValue),
                        PeriodTo = ConversionUtility.GetNullableDateTime(txtPeriodTo.EditValue),
                        Query = e.MDXQuery
                    };

                    loadManager.StartTimer();
                    OlapUtility.SetLogReport(context, Text, _windowGuid, _sessionGuid, log);
                }));
            }
        }

        private void pivotGrid_CustomFieldSort(object sender, PivotGridCustomFieldSortEventArgs e)
        {
            OlapUtility.CustomFieldSort(e, CompositeFieldsForCustomSort);
        }

        #endregion
    }
}