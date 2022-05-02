using Automation.SuperNova.PageObjects;
using Automation.SuperNova.Tracker;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Automation.SuperNova
{
    /// <summary>
    /// Process.
    /// </summary>
    public class ProcessActivity
    {
        public readonly IWebDriver webDriver;
        public readonly InputObject ipObj;

        public ProcessActivity(IWebDriver webDriver, InputObject ipObj)
        {
            this.webDriver = webDriver;
            this.ipObj = ipObj;
        }

        public void Execute()
        {
            var openOrder = new List<List<string>>();
            var openOrderShippingDetails = new List<List<string>>();
            var closeOrder = new List<List<string>>();
            new SuperNovaPage(webDriver).Execute(ipObj, openOrder, closeOrder, openOrderShippingDetails);
            var auditFileWriter = new AuditFileWriter();
            auditFileWriter.CreateTrackerAndWrite(ipObj, openOrder, closeOrder, openOrderShippingDetails);
        }
    }
}
