using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Interop.QBFC13;

namespace ImportBillToQB
{
    public partial class ImportBill : Form
    {
        public ImportBill()
        {
            
            InitializeComponent();
            //this.Text = "Import to QuickBooks";
        }
        private void button1_Click(object sender, EventArgs e)
        {
           

            //open a session to query all of the existing vendor names
            var qbVendors = new List<string>();
            
            QBSessionManager sessionManager2 = new QBSessionManager();
            sessionManager2.OpenConnection("appID", "Import To Quickbooks");
            sessionManager2.BeginSession("", ENOpenMode.omDontCare);
            IMsgSetRequest messageSet2 = sessionManager2.CreateMsgSetRequest("US", 7, 0);
            IVendorQuery vendorQuery = messageSet2.AppendVendorQueryRq();
            vendorQuery.ORVendorListQuery.VendorListFilter.ActiveStatus.SetValue(ENActiveStatus.asActiveOnly);
           
            

            try
            {
                IMsgSetResponse responseSet = sessionManager2.DoRequests(messageSet2);
                sessionManager2.EndSession();
                sessionManager2.CloseConnection();

                IResponse response;
                ENResponseType responseType;

                for (int i = 0; i < responseSet.ResponseList.Count; i++)
                {
                    response = responseSet.ResponseList.GetAt(i);
                    if (response.Detail == null) continue;
                    responseType = (ENResponseType)response.Type.GetValue();
            
                    if (responseType == ENResponseType.rtVendorQueryRs)
                    {
                        IVendorRetList vendorList = (IVendorRetList)response.Detail;
                        for (int vendorIndex = 0; vendorIndex < vendorList.Count; vendorIndex++)
                        {
                            IVendorRet vendor = (IVendorRet)vendorList.GetAt(vendorIndex);

                            if (vendor != null && vendor.CompanyName != null)
                                qbVendors.Add(vendor.CompanyName.GetValue()); //add all existing vendor names to the list qbvendors
                        }
                    }
                }
            }
            catch (System.Runtime.InteropServices.COMException comEx){ }

            //call the open file and populate our array of objects used to add bills
            List<ExtractInfo> List1 = getFile();
            
            //create two lists with existing vendors and a list of vendors from the file               
            List<string> vnames = new List<string>();
            List<string> nameList = new List<string>();
            //find which vendors from the file are not already in quickbooks
            for (int i = 0; i < List1.Count; i++)
            {
                string newVendor = List1[i].vendorName;
                nameList.Add(newVendor);
            }
            foreach (string str in nameList)
            {
                if (!qbVendors.Contains(str))
                {
                    vnames.Add(str);//the list of vendors to be added to quickbooks
                }
            }
            for (int i = 0; i < vnames.Count; i++)
            {
                string add = vnames[i];
                vendorAdd(add);//add vendors into quickbooks if not already there
            }

            populateTextBoxes(List1);
            for (int i = 0; i < List1.Count; i++)
            {
                //Open session to communicate to quickbooks
                QBSessionManager sessionManager = new QBSessionManager();
                sessionManager.OpenConnection("appID", "Create Vendor");
                sessionManager.BeginSession("", ENOpenMode.omDontCare);
                IMsgSetRequest messageSet = sessionManager.CreateMsgSetRequest("US", 13, 0);

                //Add the info to bill(vendor name, invoice num, and dates
                addBillItems(List1, messageSet, i);
                

                //tell the program to execute the changes to the bill
                IMsgSetResponse responseSet = sessionManager.DoRequests(messageSet);

                //close the session and end connection
                sessionManager.EndSession();
                sessionManager.CloseConnection();

                //report errors if any
                for (int k = 0; k < responseSet.ResponseList.Count; k++)
                {
                    IResponse response = responseSet.ResponseList.GetAt(k);
                    var code = response.StatusCode;
                    string code2 = code.ToString();

                    string code3 = response.StatusMessage;
                    if (response.StatusCode > 0)

                        MessageBox.Show(code3);
                }
            }
            
        }
        public void addBillItems(List<ExtractInfo> List1, IMsgSetRequest messageSet, int i)
        {
            // input top of bill items
            IBillAdd BillAddRequest = messageSet.AppendBillAddRq();
            BillAddRequest.VendorRef.FullName.SetValue(List1[i].vendorName);
            BillAddRequest.RefNumber.SetValue(List1[i].invoiceNumber);
            BillAddRequest.TxnDate.SetValue(DateTime.Parse(List1[i].invoiceDate));
            BillAddRequest.DueDate.SetValue(DateTime.Parse(List1[i].dueDate));
            //input expense items
            addExpense(List1, BillAddRequest, i);
            //check for repeated invoice numbers and add expense items as necessary
            for (int z = List1.Count - 1; z > i; z--)
            {
                if (List1[z].invoiceNumber != List1[i].invoiceNumber) continue;
                if (List1[z].invoiceNumber == List1[i].invoiceNumber)
                {
                    addExpense(List1, BillAddRequest, z);
                    List1.RemoveAt(z);
                }
            }
        }
        public void addExpense(List<ExtractInfo> List1, IBillAdd BillAddRequest, int i)
        {
            //add expense items (amount, memo, class, and type of expense
            IExpenseLineAdd expenseLine = BillAddRequest.ExpenseLineAddList.Append();
            expenseLine.AccountRef.FullName.SetValue(List1[i].jAccCode);
            expenseLine.Amount.SetValue(List1[i].jGrossAmount);
            expenseLine.Memo.SetValue(List1[i].lineItemDescr);
            expenseLine.ClassRef.FullName.SetValue(List1[i].reqCuston03);
        }
        public struct ExtractInfo
        {
            public string vendorName;
            public string invoiceDate;
            public string invoiceNumber;
            public double totalAmount;
            public string dueDate;
            public string jAccCode;
            public double jGrossAmount;
            public string lineItemDescr;
            public string reqCuston03;
            public string reqCustom07;
        }
        public void vendorAdd(string name)
        {
            //open session to add vendor
            QBSessionManager sessionManager = new QBSessionManager();
            sessionManager.OpenConnection("appID", "Create Vendor");
            sessionManager.BeginSession("", ENOpenMode.omDontCare);
            IMsgSetRequest messageSet = sessionManager.CreateMsgSetRequest("US", 13, 0);

            IVendorAdd vendorAddRequest = messageSet.AppendVendorAddRq();
            vendorAddRequest.Name.SetValue(name);//add vendor 
                                                 // name is the string passed from the list containing the vendors that need to be added to quickbooks

            IMsgSetResponse responseSet = sessionManager.DoRequests(messageSet);
            sessionManager.EndSession();
            sessionManager.CloseConnection();

            for (int i = 0; i < responseSet.ResponseList.Count; i++)
            {
                IResponse response = responseSet.ResponseList.GetAt(i);
                var code = response.StatusCode;
                string code2 = code.ToString();
                string code3 = response.StatusMessage;
                if (response.StatusCode > 0) { }
                // MessageBox.Show(code3);
            }
        }
        public void populateTextBoxes(List<ExtractInfo> list1)
        {
            int count = 0;
            int x1 = 375;
            int x2 = 597;
            int y = 167;
            //populate textboxes with invoice numbers and req07 fields
            for (int i = 0; i < list1.Count; i++)
            {
                TextBox tb = new TextBox();
                tb.Name = "InvoiceNum" + count;
                tb.Text = list1[count].invoiceNumber;
                tb.Location = new Point(x1, y);
                tb.Width = 150;
                
                TextBox tb2 = new TextBox();
                tb.Name = "Req07" + count;
                tb2.Text = list1[count].reqCustom07;
                tb2.Location = new Point(x2, y);
                tb2.Width = 50;
                this.Controls.Add(tb);
                this.Controls.Add(tb2);
                tb.Visible = true;
                tb2.Visible = true;
                count++;
                y = y + 26;
            }
        }
        private List<ExtractInfo> getFile()
        {
            //declare list of objects
            List<ExtractInfo> anotherList = new List<ExtractInfo>();
            ExtractInfo item1 = new ExtractInfo();
            //declare the lists that will be used
            List<String> ivDates = new List<String>();
            List<String> ivNums = new List<String>();
            List<Double> tAmounts = new List<Double>();
            List<String> dDates = new List<String>();
            List<String> jAccountCodes = new List<String>();
            List<Double> jGrossAmounts = new List<Double>();
            List<String> lineItemDescrs = new List<String>();
            List<String> reqCustoms03 = new List<String>();
            List<String> reqCustoms07 = new List<String>();
            List<String> vNames = new List<String>();
            OpenFileDialog openTxtFile = new OpenFileDialog();
            //open the file
            if (openTxtFile.ShowDialog() == DialogResult.OK)
            {
                string fileName = openTxtFile.FileName;
                List<List<string>> lines = new List<List<string>>();
                foreach (string line in File.ReadAllLines(fileName))
                {
                    var list2 = new List<string>();
                    foreach (string s in line.Split(new[] { '|' }))
                    {
                        list2.Add(s);
                    }
                    lines.Add(list2);
                }
                //Grab all data from the txt file into a 2d list
                var count = lines.Count;
                string[][] newList = new string[lines.Count][];
                for (int i = 0; i > count; i++)
                {
                    List<string> sublists = lines.ElementAt(i);
                    newList[i] = new string[sublists.Count];
                    for (int j = 0; j < sublists.Count; j++)
                    {
                        newList[i][j] = sublists.ElementAt(j);
                    }
                }
                //Populate single Lists from data in txt file from the 2d list
                for (var i = 1; i < lines.Count; i++)
                {
                    var vName = lines[i][163];
                    vNames.Add(vName);
                    var iDate = lines[i][6];
                    ivDates.Add(iDate);
                    var iNum = lines[i][5];
                    ivNums.Add(iNum);
                    double tAmount = Convert.ToDouble(lines[i][9]);
                    tAmounts.Add(tAmount);
                    var dDate = lines[i][7];
                    dDates.Add(dDate);
                    var jCode = lines[i][60];
                    jAccountCodes.Add(jCode);
                    double jAmount = Convert.ToDouble(lines[i][62]);
                    jGrossAmounts.Add(jAmount);
                    var lItemDescr = lines[i][132];
                    lineItemDescrs.Add(lItemDescr);
                    var req03 = lines[i][15];
                    string actualInput;
                    switch (req03)
                    {
                        case "111":
                            actualInput = "New Iberia";
                            reqCustoms03.Add(actualInput);
                            break;
                        case "112":
                            actualInput = "Liberty";
                            reqCustoms03.Add(actualInput);
                            break;
                        case "113":
                            actualInput = "Laurel";
                            reqCustoms03.Add(actualInput);
                            break;
                        case "114":
                            actualInput = "Houma";
                            reqCustoms03.Add(actualInput);
                            break;
                        case "115":
                            actualInput = "Thru Tubing - LA";
                            reqCustoms03.Add(actualInput);
                            break;
                        case "221":
                            actualInput = "Midland";
                            reqCustoms03.Add(actualInput);
                            break;
                        case "222":
                            actualInput = "Thru Tubing - TX";
                            reqCustoms03.Add(actualInput);
                            break;
                        case "331":
                            actualInput = "Fort Collins";
                            reqCustoms03.Add(actualInput);
                            break;
                        case "332":
                            actualInput = "Vernal";
                            reqCustoms03.Add(actualInput);
                            break;
                        case "333":
                            actualInput = "Watford City";
                            reqCustoms03.Add(actualInput);
                            break;
                        case "334":
                            actualInput = "Machine Shop";
                            reqCustoms03.Add(actualInput);
                            break;
                        case "335":
                            actualInput = "Machine Shop";
                            reqCustoms03.Add(actualInput);
                            break;
                        case "910":
                            actualInput = "Corp Ops";
                            reqCustoms03.Add(actualInput);
                            break;
                        case "920":
                            actualInput = "Corp Admin";
                            reqCustoms03.Add(actualInput);
                            break;
                        case "930":
                            actualInput = "Corp Acct";
                            reqCustoms03.Add(actualInput);
                            break;
                        case "940":
                            actualInput = "Corp Sales";
                            reqCustoms03.Add(actualInput);
                            break;
                    }
                    var req07 = lines[i][19];
                    reqCustoms07.Add(req07);
                }
                //put data from lists into object and then put object into list
                for (int i = 0; i < vNames.Count; i++)
                {
                    item1.vendorName = vNames[i];
                    item1.invoiceDate = ivDates[i];
                    item1.invoiceNumber = ivNums[i];
                    item1.reqCuston03 = reqCustoms03[i];
                    item1.lineItemDescr = lineItemDescrs[i];
                    item1.jGrossAmount = jGrossAmounts[i];
                    item1.jAccCode = jAccountCodes[i];
                    item1.dueDate = dDates[i];
                    item1.reqCustom07 = reqCustoms07[i];
                    anotherList.Add(item1);
                }
                List<ExtractInfo> List1 = anotherList.OrderBy(o => o.invoiceNumber).ToList();
                textBox1.Text = Path.GetFileName(fileName);
            }
            return anotherList;
        }
    }
}

