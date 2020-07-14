/* Title:       Blue Jay Help Desk
 * Date:        7-13-20
 * Author:      Terry Holmes
 * 
 * Description: This is used as the client for help desk */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using NewEventLogDLL;
using NewEmployeeDLL;
using HelpDeskDLL;
using PhonesDLL;
using System.Windows.Media.Converters;
using System.Security.Cryptography;

namespace BlueJayHelpDesk
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        //Setting up the classes
        WPFMessagesClass TheMessagesClass = new WPFMessagesClass();
        SendEmailClass TheSendEmailClass = new SendEmailClass();
        EventLogClass TheEventLogClass = new EventLogClass();
        EmployeeClass TheEmployeeClass = new EmployeeClass();
        HelpDeskClass TheHelpDeskClass = new HelpDeskClass();
        PhonesClass ThePhonesClass = new PhonesClass();

        //setting up the data
        FindWarehousesDataSet TheFindWarehousesDataSet = new FindWarehousesDataSet();
        FindSortedHelpDeskProblemTypeDataSet TheFindSortedHelpDeskProblemTypeDataSet = new FindSortedHelpDeskProblemTypeDataSet();
        ComboEmployeeDataSet TheComboEmployeeDataSet = new ComboEmployeeDataSet();
        FindHelpDeskTicketbyTicketMatchDateDataSet TheFindHelpDeskTicketByMatchDateDataSet = new FindHelpDeskTicketbyTicketMatchDateDataSet();
        FindPhoneExtensionByEmployeeIDDataSet TheFindPhoneExtensionByEmployeeIDDataSet = new FindPhoneExtensionByEmployeeIDDataSet();

        //setting global variables
        int gintWarehouseID;
        int gintEmployeeID;
        int gintPhoneExtension;
        int gintProblemTypeID;
        string gstrFullName;
        int gintTicketID;
        string gstrComputerName;
        string gstrUserName;
        string gstrIPAddress;
        string strOffice;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            int intCounter;
            int intNumberOfRecords;
            int intSelectedIndex = 0;
            string strLastName;

            try
            {
                //loading the warehouses
                TheFindWarehousesDataSet = TheEmployeeClass.FindWarehouses();

                cboOffice.Items.Add("Select Office");
                intNumberOfRecords = TheFindWarehousesDataSet.FindWarehouses.Rows.Count - 1;

                for (intCounter = 0; intCounter <= intNumberOfRecords; intCounter++)
                {
                    cboOffice.Items.Add(TheFindWarehousesDataSet.FindWarehouses[intCounter].FirstName);
                }

                cboOffice.SelectedIndex = 0;

                cboProblemType.Items.Add("Select Problem Type");
                TheFindSortedHelpDeskProblemTypeDataSet = TheHelpDeskClass.FindSortedHelpDeskProblemType();

                intNumberOfRecords = TheFindSortedHelpDeskProblemTypeDataSet.FindSortedHelpDeskProblemType.Rows.Count - 1;

                for (intCounter = 0; intCounter <= intNumberOfRecords; intCounter++)
                {
                    cboProblemType.Items.Add(TheFindSortedHelpDeskProblemTypeDataSet.FindSortedHelpDeskProblemType[intCounter].ProblemType);
                }

                cboProblemType.SelectedIndex = 0;

                gstrComputerName = System.Environment.MachineName.ToUpper();
                gstrUserName = System.Environment.UserName.ToUpper();
                gstrIPAddress = System.Net.Dns.GetHostEntry(System.Net.Dns.GetHostName()).AddressList.GetValue(0).ToString();

                strLastName = gstrUserName.Substring(1);

                cboSelectEmployee.Items.Add("Select Employee");

                TheComboEmployeeDataSet = TheEmployeeClass.FillEmployeeComboBox(strLastName);

                intNumberOfRecords = TheComboEmployeeDataSet.employees.Rows.Count - 1;

                if(intNumberOfRecords > -1)
                {
                    for(intCounter = 0; intCounter <= intNumberOfRecords; intCounter++)
                    {
                        cboSelectEmployee.Items.Add(TheComboEmployeeDataSet.employees[intCounter].FullName);
                    }

                    if(intNumberOfRecords == 0)
                    {
                        cboSelectEmployee.SelectedIndex = 1;
                    }
                    else
                    {
                        cboSelectEmployee.SelectedIndex = 0;
                    }
                }
                else
                {
                    strLastName = "WAREHOUSE";

                    TheComboEmployeeDataSet = TheEmployeeClass.FillEmployeeComboBox(strLastName);

                    intNumberOfRecords = TheComboEmployeeDataSet.employees.Rows.Count - 1;

                    for (intCounter = 0; intCounter <= intNumberOfRecords; intCounter++)
                    {
                        cboSelectEmployee.Items.Add(TheComboEmployeeDataSet.employees[intCounter].FullName);
                    }
                }

                if (gstrIPAddress.Contains(".0."))
                    strOffice = "CLEVELAND";
                else if (gstrIPAddress.Contains(".11."))
                    strOffice = "CBUS-GROVEPORT";
                else if (gstrIPAddress.Contains(".31."))
                    strOffice = "NASHVILLE";
                else if (gstrIPAddress.Contains(".41."))
                    strOffice = "MILWUKEE";
                else if (gstrIPAddress.Contains(".51."))
                    strOffice = "TOLEDO";
                else if (gstrIPAddress.Contains(".61."))
                    strOffice = "YOUNGSTOWN";

                intNumberOfRecords = cboOffice.Items.Count - 1;

                for (intCounter = 0; intCounter <= intNumberOfRecords; intCounter++)
                {
                    cboOffice.SelectedIndex = intCounter;

                    if(cboOffice.SelectedItem.ToString() == strOffice)
                    {
                        intSelectedIndex = cboOffice.SelectedIndex;
                    }
                }

                cboOffice.SelectedIndex = intSelectedIndex;
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Blue Jay Help Desk // Main Window // Window Loaded " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }
        }

        private void cboOffice_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            int intSelectedIndex;

            intSelectedIndex = cboOffice.SelectedIndex - 1;

            if(intSelectedIndex > -1)
            {
                gintWarehouseID = TheFindWarehousesDataSet.FindWarehouses[intSelectedIndex].EmployeeID;
            }
        }

        private void cboProblemType_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            int intSelectedIndex;

            intSelectedIndex = cboProblemType.SelectedIndex - 1;

            if(intSelectedIndex > -1)
            {
                gintProblemTypeID = TheFindSortedHelpDeskProblemTypeDataSet.FindSortedHelpDeskProblemType[intSelectedIndex].ProblemTypeID;
            }
        }

        private void cboSelectEmployee_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            int intSelectedIndex;
            int intRecordsReturned;

            try
            {
                intSelectedIndex = cboSelectEmployee.SelectedIndex - 1;

                if(intSelectedIndex > -1)
                {
                    gintEmployeeID = TheComboEmployeeDataSet.employees[intSelectedIndex].EmployeeID;
                    gstrFullName = TheComboEmployeeDataSet.employees[intSelectedIndex].FullName;

                    TheFindPhoneExtensionByEmployeeIDDataSet = ThePhonesClass.FindPhoneExtensionByEmployeeID(gintEmployeeID);

                    intRecordsReturned = TheFindPhoneExtensionByEmployeeIDDataSet.FindPhoneExtensionByEmployeeID.Rows.Count;

                    if(intRecordsReturned > 0)
                    {
                        gintPhoneExtension = TheFindPhoneExtensionByEmployeeIDDataSet.FindPhoneExtensionByEmployeeID[0].Extension;
                    }
                    else
                    {
                        gintPhoneExtension = 0;
                    }
                }
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Blue Jay Help Desk // Main Window // Employee Combo Box " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }
        }

        private void btnSubmit_Click(object sender, RoutedEventArgs e)
        {
            bool blnFatalError = false;
            string strErrorMessage = "";
            DateTime datTicketDate = DateTime.Now;
            string strRepotedProblem;
            string strEmailAddress = "itadmin@bluejaycommunications.com";
            string strHeader;
            string strMessage;
            

            try
            {
                //data valication
                if (cboOffice.SelectedIndex < 1)
                {
                    blnFatalError = true;
                    strErrorMessage += "The Office Was Not Selected\n";
                }
                if (cboSelectEmployee.SelectedIndex < 1)
                {
                    blnFatalError = true;
                    strErrorMessage += "The Employee Was Not Selected\n";
                }
                if (cboProblemType.SelectedIndex < 1)
                {
                    blnFatalError = true;
                    strErrorMessage += "The Problem Type Was Not Selected";
                }
                strRepotedProblem = txtProblemNotes.Text;
                if (strRepotedProblem.Length < 10)
                {
                    blnFatalError = true;
                    strErrorMessage += "The Reported Problem is not Long Enough\n";
                }
                if(blnFatalError == true)
                {
                    TheMessagesClass.ErrorMessage(strErrorMessage);
                    return;
                }

                //inserting ticket
                blnFatalError = TheHelpDeskClass.InsertHelpDeskTicket(datTicketDate, gstrComputerName, gstrUserName, gstrIPAddress, gintWarehouseID,gintProblemTypeID, strRepotedProblem, gintEmployeeID);

                if (blnFatalError == true)
                    throw new Exception();

                TheFindHelpDeskTicketByMatchDateDataSet = TheHelpDeskClass.FindHelpDeskTicketByTicketDateMatch(datTicketDate);

                gintTicketID = TheFindHelpDeskTicketByMatchDateDataSet.FindHelpDeskTicketByTicketDateMatch[0].TicketID;

                blnFatalError = TheHelpDeskClass.InsertHelpDeskTicketUpdate(gintTicketID, gintEmployeeID, "TICKET CREATED");

                if (blnFatalError == true)
                    throw new Exception();

                strHeader = gstrFullName + " Has Submitted a Help Desk Ticket - Do Not Reply";
                strMessage = "<h1>" + gstrFullName + " Has Submitted a Help Desk Ticket - Do Not Reply</h1>";
                strMessage += "<h3> Ticket ID " + Convert.ToString(gintTicketID) + "</h3>";
                strMessage += "<h3> They have Reported The Following Problem </h3>";
                strMessage += "<h3>" + strRepotedProblem + "</h3>";
                strMessage += "<h3> They can be reached at Extension " + Convert.ToString(gintPhoneExtension) + "</h3>";
                strMessage += "<h3> Computer Name " + gstrComputerName + " User Name " + gstrUserName + " IP Address " + gstrIPAddress + "<h3>";

                blnFatalError = TheSendEmailClass.SendEmail(strEmailAddress, strHeader, strMessage);

                if (blnFatalError == false)
                    throw new Exception();

                TheMessagesClass.InformationMessage("Help Desk Ticket Number " + Convert.ToString(gintTicketID) + " Has Been Created");
                
            }
            catch(Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Blue Jay Help Desk // Main Window // Submit Button " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }
        }
    }
}
