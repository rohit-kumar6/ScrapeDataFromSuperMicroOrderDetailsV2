using Automation.Core.Excel;
using Serilog;
using System;
using System.Collections.Generic;

namespace Automation.SuperNova.Tracker
{
    /// <summary>
    /// Class responsible for writing RO Details to audit file tracker.
    /// </summary>
    public class AuditFileWriter
    {
        /// <summary>
        /// Create Tracker and Write.
        /// </summary>
        /// <param name="inputObject">Input Object.</param>
        public void CreateTrackerAndWrite(
            InputObject inputObject, 
            List<List<string>> openOrder, 
            List<List<string>> closeOrder,
            List<List<string>> openOrderShippingDetails)
        {
            ExcelOperations trackerExcel = new ExcelOperations();
            try
            {
                string filename = $@"{inputObject.outputPath}\Ouput_{DateTime.Now:yyyy_MM_dd_hh_mm_ss}.xlsx";
                _ = new ExcelOperations(filename);
                trackerExcel = new ExcelOperations(filename, 1);

                trackerExcel.RenameActiveSheet("Open Order");
                var headers = GetOpenOrderHeaders();
                trackerExcel.AddHeaders(headers);
                WriteOrderData(trackerExcel, openOrder);

                trackerExcel.CreateNewSheet("Close Order");
                trackerExcel.ActivateSheet("Close Order");
                headers = GetCloseOrderHeaders();
                trackerExcel.AddHeaders(headers);
                WriteOrderData(trackerExcel, closeOrder);

                trackerExcel.CreateNewSheet("Open Order Shipping Details");
                trackerExcel.ActivateSheet("Open Order Shipping Details");
                headers = GetOpenOrderShippingHeaders();
                trackerExcel.AddHeaders(headers);
                WriteOrderData(trackerExcel, openOrderShippingDetails);
            }
            catch (Exception ex)
            {
                Log.Information("Error in writing audit file.");
                Log.Error(ex.Message + " " +  ex.StackTrace);
            }
            finally
            {
                trackerExcel.Save();
                trackerExcel.Close();
            }
        }

        private static string[,] GetCloseOrderHeaders()
        {
            var headers = new string[,]
            {
                     {
                        "Sold To ID", "Sales Order", "Order Date", "Customer PO", "Assembly Type",
                        "Order Status", "Sold-To", "Ship-To", "Line No.",
                        "Item Number", "Description", "QTY Ordered", "QTY Shipped",
                        "B/O QTY", "Unit Price", "Extended Price"
                     },
            };

            return headers;
        }

        private static string[,] GetOpenOrderShippingHeaders()
        {
            var headers = new string[,]
            {
                     {
                        "Sold To ID", "Sales Order", "PICKUNIQ", "ORDUNIQ", "Status",
                        "VIADESC", "Comment", "SDATE", "Invoice Number", "Tracking Number",
                        "Order Line", "QTY Ordered", "QTY Shipped", "Item",
                        "Category", "Description", "Timestamps"
                     },
            };

            return headers;
        }

        private static string[,] GetOpenOrderHeaders()
        {
            var headers = new string[,]
            {
                     {
                        "Sold To ID", "Sales Order", "Customer PO", "Ship To Party",
                        "Ship To Country", "Created Time", "Assembly Type",
                        "Order Status", "Sold-To", "Ship-To", "ESD", "Message",
                        "Line No.", "Item Number", "Description", "QTY Ordered",
                        "QTY Shipped", "B/O QTY", "Unit Price", "Extended Price"
                     },
            };

            return headers;
        }

        private void WriteOrderData(ExcelOperations trackerExcel, List<List<string>> orderList)
        {
            int row = 1;
            int col = 1;
            foreach (var order in orderList)
            {
                row++;
                foreach (var item in order)
                {
                    trackerExcel.WriteStringToCell(row, col++, item.ToString());
                }

                col = 1;
            }

            trackerExcel.AutoFit();
        }
    }
}
