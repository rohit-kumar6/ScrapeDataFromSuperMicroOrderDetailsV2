namespace Automation.SuperNova.PageObjects
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading;
    using Automation.Core;
    using Automation.Core.Web;
    using OpenQA.Selenium;
    using Serilog;

    /// <summary>
    /// Page Object for SuperNovaPage.
    /// </summary>
    [LocatorResource(resourceName: "SuperNovaPage", resourceType: typeof(ResourceFile))]
    public class SuperNovaPage : BasePage
    {
        [Locator("ORDER_TYPE")]
        private readonly UIElement _orderType;
        
        [Locator("FROM_DATE")]
        private readonly UIElement _fromDate;
        
        [Locator("TO_DATE")]
        private readonly UIElement _toDate;
        
        [Locator("CUSTOMER_ID")]
        private readonly UIElement _customerId;
        
        [Locator("SEARCH_BUTTON")]
        private readonly UIElement _searchButton;
        
        [Locator("ORDER_TABLE")]
        private readonly UIElement _orderTable;
        
        [Locator("OPEN_ORDER_TABLE_ROW_COUNT")]
        private readonly UIElement _openOrderTableRowCount;
        
        [Locator("OPEN_ORDER_SOLD_TO_ID")]
        private readonly UIElement _openOrderSoldToId;
        
        [Locator("OPEN_ORDER_SALES_ORDER")]
        private readonly UIElement _openOrderSalesOrder;

        [Locator("openOrderCustomerPO")]
        private readonly UIElement _openOrderCustomerPO;

        [Locator("openOrderShipToParty")]
        private readonly UIElement _openOrderShipToParty;
        
        [Locator("openOrderShipToCountry")]
        private readonly UIElement _openOrderShipToCountry;
        
        [Locator("openOrderCreatedTime")]
        private readonly UIElement _openOrderCreatedTime;
        
        [Locator("openOrderOrderDetailsButton")]
        private readonly UIElement _openOrderDetailsButton;
        
        [Locator("openOrderAssemblyType")]
        private readonly UIElement _openOrderAssemblyType;
        
        [Locator("openOrderOrderStatus")]
        private readonly UIElement _openOrderOrderStatus;
        
        [Locator("SHIPPING_DETAILS_BUTTON")]
        private readonly UIElement _shippingDetailsButton;
        
        [Locator("NO_SHIPPING_DETAILS_AVAILABLE")]
        private readonly UIElement _NoshippingDetailsButton;
        
        [Locator("SHIPPING_DETAILS_VALUE")]
        private readonly UIElement _shippingDetailsValue;
        
        [Locator("SHIPPING_DETAILS_MORE_BUTTON")]
        private readonly UIElement _shippingDetailsMoreButton;
        
        [Locator("SHIPPING_DETAILS_ORDER_ROW")]
        private readonly UIElement _shippingDetailsOrderRow;
        
        [Locator("SHIPPING_DETAILS_ORDER_VALUE")]
        private readonly UIElement _shippingDetailsOrderValue;
        
        [Locator("soldToAddress")]
        private readonly UIElement _soldToAddress;
        
        [Locator("shipToAddress")]
        private readonly UIElement _shipToAddress;
        
        [Locator("message")]
        private readonly UIElement _message;
        
        [Locator("esd")]
        private readonly UIElement _esd;
        
        [Locator("closeOrderTableRowCount")]
        private readonly UIElement _closeOrderTableRowCount;
        
        [Locator("closeOrderSalesOrder")]
        private readonly UIElement _closeOrderSalesOrder;
        
        [Locator("closeOrderOrderData")]
        private readonly UIElement _closeOrderOrderData;
        
        [Locator("closeOrderCustomerPO")]
        private readonly UIElement _closeOrderCustomerPO;
        
        [Locator("closeOrderAssemblyType")]
        private readonly UIElement _closeOrderAssemblyType;
        
        [Locator("closeOrderOrderDetailsButton")]
        private readonly UIElement _closeOrderOrderDetailsButton;
        
        [Locator("closeOrderOrderStatus")]
        private readonly UIElement _closeOrderOrderStatus;
        
        [Locator("orderItem")]
        private readonly UIElement _orderItem;
        
        [Locator("orderItemTable")]
        private readonly UIElement _orderItemTable;
        
        [Locator("orderItemTableNoDataAvailable")]
        private readonly UIElement _orderItemTableNoDataAvailable;
        
        [Locator("orderItemHeaders")]
        private readonly UIElement _orderItemHeaders;
        
        [Locator("orderItemRows")]
        private readonly UIElement _orderItemRows;
        
        [Locator("orderItemRowsValue")]
        private readonly UIElement _orderItemRowsValue;
        
        [Locator("orderItemHeadersValue")]
        private readonly UIElement _orderItemHeadersValue;
        
        [Locator("ORDER_ITEM_EXPANDED")]
        private readonly UIElement _orderItemExpanded;
        
        [Locator("VIEW_DETAILS_OPENORDER_CHECKBOX")]
        private readonly UIElement _viewDetailsOpenOrderCheckbox;
        
        [Locator("VIEW_DETAILS_CLOSEORDER_CHECKBOX")]
        private readonly UIElement _viewDetailsCloseOrderCheckbox;
        
        [Locator("PAGE_COUNT")]
        private readonly UIElement _pageCount;
        
        [Locator("SELECT_PAGE")]
        private readonly UIElement _selectPage;
        
        [Locator("SDATE")]
        private readonly UIElement _sdate;

        private readonly IWebDriver _webDriver;
        private readonly RetryExecutor _retryExecutor;

        /// <summary>
        /// Initializes a new instance of the <see cref="SuperNovaPage"/> class.
        /// </summary>
        /// <param name="webDriver">WebDriver instance.</param>
        public SuperNovaPage(IWebDriver webDriver)
            : base(webDriver, 120)
        {
            Initialize();
            _webDriver = webDriver;
            _retryExecutor = new RetryExecutor(new RetryStrategy(3, TimeSpan.FromSeconds(15)));
        }

        public void Execute(
            InputObject ipObj,
            List<List<string>> openOrder,
            List<List<string>> closeOrder,
            List<List<string>> openOrderShippingDetails)
        {
            foreach (var orderType in ipObj.orderTypeList)
            {
                Log.Information($"Processing order type {orderType}");
                _retryExecutor.Retry(() =>
                {
                    OpenWebPage(new Uri("https://customerportal.supermicro.com/SO/CustOrders.aspx"));
                    _orderType.WaitForElementToBeVisible(300);
                    _orderType.SelectByText(orderType);
                    _fromDate.SetText(ipObj.startDate);
                    _toDate.SetText(ipObj.endDate);
                });
                
                Log.Information($"Start date {ipObj.startDate} - End Date {ipObj.endDate}");

                foreach (var customerId in ipObj.customerIdList)
                {
                    Log.Information($"Processing customer id {customerId}");

                    _retryExecutor.Retry(() =>
                    {
                        _customerId.SelectByValue(customerId);
                        _searchButton.Click();
                    });

                    GetDataOfCurrentPage(orderType, openOrder, closeOrder, customerId, openOrderShippingDetails);
                    int curPageNo = 2;
                    bool dotClicked = false;
                    bool endPageReached = false;

                    while (true)
                    {
                        if (!_pageCount.IsVisible(2))
                        {
                            break;
                        }

                        if (endPageReached)
                        {
                            break;
                        }

                        int pageCount = _pageCount.GetSize();
                        for (int page = 2; page <= pageCount; page++)
                        {
                            string pageText = GetDynamicUIElement(_selectPage, page).GetText();
                            if (page == pageCount && pageText.Equals("..."))
                            {
                                GetDynamicUIElement(_selectPage, page).Click();
                                dotClicked = true;
                                break;
                            }

                            if (pageText.Equals(curPageNo.ToString()))
                            {
                                if(!dotClicked)
                                {
                                    GetDynamicUIElement(_selectPage, page).Click();
                                }

                                dotClicked = false;
                                Thread.Sleep(5000);
                                Log.Information($"Processing page no {curPageNo}");
                                GetDataOfCurrentPage(orderType, openOrder, closeOrder, customerId, openOrderShippingDetails);
                                curPageNo += 1;
                                if (page == pageCount)
                                {
                                    endPageReached = true;
                                }
                            }
                        }
                    }

                    Thread.Sleep(2000);
                }

                Thread.Sleep(2000);
            }
        }

        private void GetDataOfCurrentPage(
            string orderType,
            List<List<string>> openOrder,
            List<List<string>> closeOrder,
            string customerId,
            List<List<string>> openOrderShippingDetails)
        {
            if (orderType.Equals("Open Order"))
            {
                GetOpenOrderDetails(openOrder, customerId, openOrderShippingDetails);
            }
            else
            {
                GetClosedOrderDetails(closeOrder, customerId);
            }
        }

        private void GetClosedOrderDetails(List<List<string>> closeOrder, string customerId)
        {
            int rowCount = _closeOrderTableRowCount.GetSize();
            if (_pageCount.IsVisible(1))
            {
                rowCount -= 2;
            }

            for (int i = 2; i <= rowCount; i++)
            {
                var closeOrderTemp = closeOrder.ToList();
                _retryExecutor.Retry(() =>
                {
                    var tempList = new List<string>();
                    tempList.Add(customerId);
                    var salesOrder = GetDynamicUIElement(_closeOrderSalesOrder, i);
                    salesOrder.ScrollIntoView();
                    tempList.Add(salesOrder.GetText());
                    tempList.Add(GetDynamicUIElement(_closeOrderOrderData, i).GetText());
                    tempList.Add(GetDynamicUIElement(_closeOrderCustomerPO, i).GetText());
                    tempList.Add(GetDynamicUIElement(_closeOrderAssemblyType, i).GetText());
                    GetDynamicUIElement(_closeOrderOrderDetailsButton, i).Click();
                    _closeOrderOrderStatus.WaitForElementToBeVisible(120);
                    tempList.Add(_closeOrderOrderStatus.GetText());
                    tempList.Add(GetSoldToAddress());
                    tempList.Add(GetShipToAddress());
                    tempList.Add(_sdate.GetText());
                    GetOrderItems(tempList, closeOrder, true);
                }, () => {
                    closeOrder = closeOrderTemp.ToList();
                });
            }
        }

        private void GetOpenOrderDetails(
            List<List<string>> openOrder, string customerId, List<List<string>> openOrderShippingDetails)
        {
            int rowCount = _openOrderTableRowCount.GetSize();
            if (_pageCount.IsVisible(10))
            {
                rowCount -= 1;
            }

            for (int i = 2; i <= rowCount; i++)
            {
                var openOrderTemp = openOrder.ToList();
                var openOrderShippingDetailsTemp = openOrderShippingDetails.ToList();
                _retryExecutor.Retry(() =>
                {
                    var tempList = new List<string>();
                    tempList.Add(customerId);
                    var soldToId = GetDynamicUIElement(_openOrderSoldToId, i);
                    soldToId.ScrollIntoView();
                    tempList.Add(soldToId.GetText());
                    string salesOrder = GetDynamicUIElement(_openOrderSalesOrder, i).GetText();
                    tempList.Add(salesOrder);
                    tempList.Add(GetDynamicUIElement(_openOrderCustomerPO, i).GetText());
                    tempList.Add(GetDynamicUIElement(_openOrderShipToParty, i).GetText());
                    tempList.Add(GetDynamicUIElement(_openOrderShipToCountry, i).GetText());
                    tempList.Add(GetDynamicUIElement(_openOrderCreatedTime, i).GetText());
                    var openOrderDetailsButton = GetDynamicUIElement(_openOrderDetailsButton, i);
                    if (openOrderDetailsButton.IsVisible(2))
                    {
                        openOrderDetailsButton.Click();
                    }

                    _openOrderAssemblyType.WaitForElementToBeVisible(120);
                    tempList.Add(_openOrderAssemblyType.GetText());
                    tempList.Add(_openOrderOrderStatus.GetText());
                    tempList.Add(GetSoldToAddress());
                    tempList.Add(GetShipToAddress());
                    tempList.Add(_esd.GetText());
                    tempList.Add(_message.GetText());
                    GetOrderItems(tempList, openOrder, false);
                    GetShippingDetails(customerId, salesOrder, openOrderShippingDetails);
                }, () => {
                    openOrder = openOrderTemp.ToList();
                    openOrderShippingDetails = openOrderShippingDetailsTemp.ToList();
                });
            }
        }

        private void GetShippingDetails(string customerId, string salesOrder, List<List<string>> openOrderShippingDetails)
        {
            _shippingDetailsButton.ScrollIntoView();
            _shippingDetailsButton.WaitAndClick(10);
            if (_NoshippingDetailsButton.IsVisible(5))
            {
                return;
            }

            var tempList = new List<string>();
            tempList.Add(customerId);
            tempList.Add(salesOrder);
            tempList.Add(GetDynamicUIElement(_shippingDetailsValue, 1).GetText());
            tempList.Add(GetDynamicUIElement(_shippingDetailsValue, 2).GetText());
            tempList.Add(GetDynamicUIElement(_shippingDetailsValue, 3).GetText());
            tempList.Add(GetDynamicUIElement(_shippingDetailsValue, 4).GetText());
            tempList.Add(GetDynamicUIElement(_shippingDetailsValue, 5).GetText());
            tempList.Add(GetDynamicUIElement(_shippingDetailsValue, 6).GetText());
            tempList.Add(GetDynamicUIElement(_shippingDetailsValue, 7).GetText());
            tempList.Add(GetDynamicUIElement(_shippingDetailsValue, 8).GetText());
            _shippingDetailsMoreButton.ScrollIntoView();
            _shippingDetailsMoreButton.Click();
            Thread.Sleep(2000);

            int orderRow = _shippingDetailsOrderRow.GetSize();
            for (int i = 2; i <= orderRow; i++)
            {
                var tempListCopy = tempList.ToList();
                tempListCopy.Add(GetDynamicUIElement(_shippingDetailsOrderValue, i, 1).GetText());
                tempListCopy.Add(GetDynamicUIElement(_shippingDetailsOrderValue, i, 2).GetText());
                tempListCopy.Add(GetDynamicUIElement(_shippingDetailsOrderValue, i, 3).GetText());
                tempListCopy.Add(GetDynamicUIElement(_shippingDetailsOrderValue, i, 4).GetText());
                tempListCopy.Add(GetDynamicUIElement(_shippingDetailsOrderValue, i, 5).GetText());
                tempListCopy.Add(GetDynamicUIElement(_shippingDetailsOrderValue, i, 6).GetText());
                tempListCopy.Add(GetDynamicUIElement(_shippingDetailsOrderValue, i, 7).GetText());
                openOrderShippingDetails.Add(tempListCopy);
            }
        }

        private string GetShipToAddress()
        {
            string shipToAddress = GetDynamicUIElement(_shipToAddress, 1).GetText();
            shipToAddress += " " + GetDynamicUIElement(_shipToAddress, 2).GetText();
            return shipToAddress;
        }

        private string GetSoldToAddress()
        {
            string soldToAddress = GetDynamicUIElement(_soldToAddress, 1).GetText();
            soldToAddress += " " + GetDynamicUIElement(_soldToAddress, 2).GetText();
            return soldToAddress;
        }

        private UIElement GetDynamicUIElement(UIElement uiElement, int row)
        {
            return new UIElement(
                       _webDriver,
                       uiElement.LocatorType,
                       string.Format(uiElement.LocatorValue, row),
                       10);
        }
        
        private UIElement GetDynamicUIElement(UIElement uiElement, int row, int col)
        {
            return new UIElement(
                       _webDriver,
                       uiElement.LocatorType,
                       string.Format(uiElement.LocatorValue, row, col),
                       10);
        }
        
        private void GetOrderItems(List<string> tempList, List<List<string>> orderDetailsList, bool closedOrder)
        {
            Thread.Sleep(2000);
            _orderItem.ScrollIntoView();
            _orderItem.WaitAndClick(60);
            _orderItemTable.WaitForElementToBeVisible(120);
            _orderItemTable.ScrollIntoView();

            try
            {
                if (_orderItemTableNoDataAvailable.IsVisible(2))
                {
                    orderDetailsList.Add(tempList.ToList());
                    return;
                }

                if (!_orderItemHeaders.IsVisible(5))
                {
                    throw new Exception("Not opened");
                }

                _orderItemHeaders.GetSize();
            }
            catch (Exception ex)
            {
                _orderItem.ScrollIntoView();
                _orderItem.WaitAndClick(60);
                _orderItemTable.WaitForElementToBeVisible(120);
                _orderItemTable.ScrollIntoView();
            }

            if (closedOrder && _viewDetailsCloseOrderCheckbox.IsVisible(5))
            {
                _viewDetailsCloseOrderCheckbox.Check();
            }
            else if (!closedOrder && _viewDetailsOpenOrderCheckbox.IsVisible(5))
            {
                _viewDetailsOpenOrderCheckbox.Check();
            }

            Thread.Sleep(2000);
            if (!_orderItemTable.IsVisible(2))
            {
                _orderItem.ScrollIntoView();
                _orderItem.WaitAndClick(60);
                _orderItemTable.WaitForElementToBeVisible(120);
                _orderItemTable.ScrollIntoView();
            }

            if (_orderItemTableNoDataAvailable.IsVisible(2))
            {
                orderDetailsList.Add(tempList.ToList());
                return;
            }

            int headersCount = _orderItemHeaders.GetSize();
            if(headersCount == 0)
            {
                orderDetailsList.Add(tempList.ToList());
                return;
            }

            var headersDict = new Dictionary<string, int>();
            for (int i = 1; i <= headersCount; i++)
            {
                headersDict.Add(GetDynamicUIElement(_orderItemHeadersValue, i).GetText(), i);
            }

            int rowsCount = _orderItemRows.GetSize();
            for (int row = 2; row <= rowsCount; row++)
            {
                var tempListCopy = tempList.ToList();
                tempListCopy.Add(headersDict.ContainsKey("Line No.") ?
                    GetDynamicUIElement(_orderItemRowsValue, row, headersDict["Line No."]).GetText() : string.Empty);
                tempListCopy.Add(headersDict.ContainsKey("Item Number") ?
                    GetDynamicUIElement(_orderItemRowsValue, row, headersDict["Item Number"]).GetText() : string.Empty);
                tempListCopy.Add(headersDict.ContainsKey("Description") ?
                    GetDynamicUIElement(_orderItemRowsValue, row, headersDict["Description"]).GetText() : string.Empty);
                tempListCopy.Add(headersDict.ContainsKey("QTY Ordered") ?
                    GetDynamicUIElement(_orderItemRowsValue, row, headersDict["QTY Ordered"]).GetText() : string.Empty);
                tempListCopy.Add(headersDict.ContainsKey("QTY Shipped") ?
                    GetDynamicUIElement(_orderItemRowsValue, row, headersDict["QTY Shipped"]).GetText() : string.Empty);
                tempListCopy.Add(headersDict.ContainsKey("B/O QTY") ?
                    GetDynamicUIElement(_orderItemRowsValue, row, headersDict["B/O QTY"]).GetText() : string.Empty);
                tempListCopy.Add(headersDict.ContainsKey("Unit Price") ?
                    GetDynamicUIElement(_orderItemRowsValue, row, headersDict["Unit Price"]).GetText() : string.Empty);
                tempListCopy.Add(headersDict.ContainsKey("Extended Price") ?
                    GetDynamicUIElement(_orderItemRowsValue, row, headersDict["Extended Price"]).GetText() : string.Empty);
                orderDetailsList.Add(tempListCopy);
            }
        }
    }
}