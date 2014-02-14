/**
 * Form1.cs
 * Author: Joshua Parker
 * ID: 1233877
 * UPI: jpar390
 * 
 * The program loads the database data and displays it to the user.
 * Lets the user modify and save the changes they made to the database.
 */

using System;
using System.Data;
using System.Windows.Forms;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace CS280A2
{
    public partial class Form1 : Form
    {
        /**
         * Initialize the form
         */
        public Form1()
        {
            InitializeComponent();
            this.orderDateDataGridViewTextBoxColumn.DataGridView.DataError += new DataGridViewDataErrorEventHandler(ordersDataGridView_DataError);
            this.shipperIDDataGridViewTextBoxColumn.DataGridView.DataError += new DataGridViewDataErrorEventHandler(shippersDataGridView_DataError);
            this.customerIDDataGridViewTextBoxColumn.DataGridView.DataError += new DataGridViewDataErrorEventHandler(customersDataGridView_DataError);
            SetControls(false);
        }

        /**
         * Following 3 methods are to catch data grid view exceptions for each table
         */
        private void ordersDataGridView_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            if (e.Exception != null && e.Context == DataGridViewDataErrorContexts.Commit)
            {
                endEditOrders(sender, e);
            }
        }

        // same as above
        private void shippersDataGridView_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            if (e.Exception != null && e.Context == DataGridViewDataErrorContexts.Commit)
            {
                this.Validate();
                this.shippersBindingSource.EndEdit();
                a2DataSet.Shippers.ColumnChanging += Column_Changing;
            }
        }

        // same as above
        private void customersDataGridView_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            if (e.Exception != null && e.Context == DataGridViewDataErrorContexts.Commit)
            {
                this.Validate();
                this.customersBindingSource.EndEdit();
                a2DataSet.Customers.ColumnChanging += Column_Changing;
            }
        }

        /**
         * Checks for cloumn changes
         */
        private void Form1_Load(object sender, EventArgs e)
        {
            a2DataSet.Customers.ColumnChanging += Column_Changing;
            a2DataSet.Shippers.ColumnChanging += Column_Changing;
            a2DataSet.Orders.ColumnChanging += Column_Changing;
        }

        /**
         * Called once a text box has been modified
         */
        private void endEditOrders(object sender, EventArgs e)
        {
            try
            {
                Form1_Load(sender, e);
                this.Validate();
                this.customersBindingSource.EndEdit();
                this.shippersBindingSource.EndEdit();
                this.ordersBindingSource.EndEdit();
            }
            catch (Exception)
            {
                MessageBox.Show("Sorry that input is invalid, changes will be rolled back");
                this.ordersBindingSource.CancelEdit();
                this.customersBindingSource.CancelEdit();
                this.shippersBindingSource.CancelEdit();
            }
        }

        /**
         * Enables or disables the menu items / program elements
         * Cannot use tables until a database is loaded
         */
        private void SetControls(bool value)
        {
            //TODO: Make sure menu items are correctly blocked out
            loadToolStripMenuItem.Enabled = !value;
            exportAsToolStripMenuItem.Enabled = value;
            saveToolStripMenuItem.Enabled = value;
            orderIDTextBox.Enabled = value;
            comboBox1.Enabled = value;
            trackingIDTextBox.Enabled = value;
            shippedDateTextBox.Enabled = value;
            button1.Enabled = value;
            dataGridView1.Enabled = value;
            dataGridView2.Enabled = value;
            dataGridView3.Enabled = value;
            bindingNavigatorAddNewItem.Enabled = value;
            bindingNavigatorAddNewItem1.Enabled = value;
            bindingNavigatorAddNewItem2.Enabled = value;
        }

        /**
         * Reads in a UDL file to connect to the database defined in the file
         */
        private void loadToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                using (OpenFileDialog dlgOpenFile = new OpenFileDialog())
                {
                    dlgOpenFile.Filter = "UDL Files(*.udl)|*.udl";
                    if (dlgOpenFile.ShowDialog() == DialogResult.OK)
                    {
                        //Change the connectionstring only, don't create a new connection object
                        customersTableAdapter.Connection.ConnectionString = "File Name=" + dlgOpenFile.FileName;
                        ordersTableAdapter.Connection.ConnectionString = "File Name=" + dlgOpenFile.FileName;
                        shippersTableAdapter.Connection.ConnectionString = "File Name=" + dlgOpenFile.FileName;
                        //load tables - parents then child
                        this.customersTableAdapter.Fill(this.a2DataSet.Customers);
                        this.shippersTableAdapter.Fill(this.a2DataSet.Shippers);
                        this.ordersTableAdapter.Fill(this.a2DataSet.Orders);

                        //uses the IsInitialized result as a boolean to enables controls if table is initalized
                        SetControls(this.customersTableAdapter.GetData().IsInitialized);

                        MessageBox.Show("Loading done");
                    }
                }
            }
            catch (System.Data.OleDb.OleDbException)
            {
                MessageBox.Show("Error: Please check your udl file, as it does not point to a database (*.accdb) file", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);   
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /**
         * Displays the about box to the user
         */
        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AboutBox1 aboutForm = new AboutBox1();
            aboutForm.ShowDialog();
        }

        /**
         * Generates the text to output to the customer when sending a message
         * If no tracking id is present the file is generated without one
         */
        private string customerMessage(string id, string dateOrder, string dateSent, string trackingNo, string sentBy)
        {
            string mess = "Hi,\r\n\r\n"
                + "The tracking number and details for your order are below:\r\n"
                + "ID: " + id + "\r\n"
                + "Order date: " + Convert.ToDateTime(dateOrder).ToShortDateString() + "\r\n"
                + "Tracking: " + trackingNo + "\r\n"
                + "Sent on " + Convert.ToDateTime(dateSent).ToShortDateString() + " by " + sentBy + "\r\n"
                + "\r\nRegards\r\n"
                + new AboutBox1().AssemblyCompany.ToString();
            return mess;
        }

        /**
         * The send button to "send" a message to the customer regarding their order
         */
        private void button1_Click_1(object sender, EventArgs e)
        {
            StreamWriter sw = null;
            DataRow dr = ((DataRowView)ordersBindingSource.Current).Row;
            string orderOutput = dr["OrderID"].ToString();

            try
            {
                orderOutput += ".txt";
                if (File.Exists(orderOutput))
                    File.Delete(orderOutput);
                if (dr["OrderID"].ToString() == "0")
                    MessageBox.Show("Order ID cannot be 0", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                else if (dr["OrderID"].ToString().Length < 1 || dr["OrderDate"].ToString().Length < 1 || dr["ShippedDate"].ToString().Length < 1 || dr["ShippedDate"].ToString().Length < 1)
                    MessageBox.Show("Missing Value please check that order id is not 0 and order date, shipped date and company name are all not empty", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                else if (dr["TrackingID"].ToString().Length <= 0)
                    MessageBox.Show("Tracking information cannot be output, no Tracking ID entered", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                else
                {
                    sw = new StreamWriter(orderOutput, true);
                    sw.WriteLine(customerMessage(orderOutput, dr["OrderDate"].ToString(),
                                                 dr["ShippedDate"].ToString(), dr["TrackingID"].ToString(),
                                                 a2DataSet.Shippers.Rows[int.Parse(dr["ShipVia"].ToString()) - 1][1].ToString()));
                    MessageBox.Show("A File has been written to the following folder:\n" + Directory.GetCurrentDirectory() + "\\" + orderOutput);
                }
            }
            catch (Exception ex)
            {
                // delete any files that were created before the error was encountered
                if (File.Exists(orderOutput))
                {
                    if (sw != null) sw.Close();
                    File.Delete(orderOutput);
                }
                Console.WriteLine("The following error occured:\n" + ex.ToString());
            }
            finally
            {
                if (sw != null)
                    sw.Close();
            }
        }

        /**
         * To give an undefined value for any file where a field was left empty
         */
        private string checkForEmpty(string given)
        {
            if (given.Length < 1)
                return "Undefined";
            return given;
        }

        /**
         * Outputs the unshipped orders data to an excel file
         */
        private void excelFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string saveFileAs = "report.xlsx";
            String[] rowHeadings = new String[] { "Order ID", "Contact Name", "Date Ordered" };
            exportToExcel(saveFileAs, rowHeadings, true);
        }

        /**
         * Writes an excel file of orders for the user to read
         */
        private void exportToExcel(string saveFileAs, string[] rowHeadings, bool unshippedOrders)
        {
            Excel.Application xlApp = null;
            Excel.Workbook xlBook = null;
            Excel.Worksheet xlSheet1 = null;
            Excel.Range xlRange = null;
            try
            {
                xlApp = new Excel.Application();
                // hide excel and alerts while generating the report
                xlApp.Visible = false;
                xlApp.DisplayAlerts = false;

                // create a new workbook
                xlBook = xlApp.Workbooks.Add(Type.Missing);
                xlSheet1 = (Excel.Worksheet)xlBook.Sheets[1];
                xlRange = xlSheet1.get_Range("A1", Type.Missing);
                //settings depending on if wanting shipped or unshipped orders
                if (unshippedOrders)
                {
                    xlSheet1.Name = "Unshipped Orders Report";
                    xlSheet1.get_Range("A1", "C1").Merge(Type.Missing);
                    xlRange.set_Value(Type.Missing, "Unshipped Orders Report");
                }
                else
                {
                    xlSheet1.Name = "Shipped Orders Report";
                    xlSheet1.get_Range("A1", "D1").Merge(Type.Missing);
                    xlRange.set_Value(Type.Missing, "Shipped Orders Report");
                }
                // make a title
                xlRange.Font.Name = "Cambria";
                xlRange.Font.Bold = true;
                xlRange.Font.Size = 12;
                xlRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                
                //heading
                if(unshippedOrders)
                    xlRange = xlSheet1.get_Range("A2:C2", Type.Missing);
                else
                    xlRange = xlSheet1.get_Range("A2:D2", Type.Missing);
                xlRange.set_Value(Type.Missing, rowHeadings);

                //Warning for shipped orders
                if (!unshippedOrders)
                    MessageBox.Show("Please note this may take a few moments to complete", "WARNING", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                //Need to loop through customer data and make a list of shipped/unshipped orders
                var tmp = this.a2DataSet.Orders;
                var tmp2 = this.a2DataSet.Customers;
                int i = 3;
                foreach (var t in tmp)
                {
                    try
                    {
                        if ((t.IsShippedDateNull() && unshippedOrders) || (!t.IsShippedDateNull() && !unshippedOrders))
                            foreach (var t2 in tmp2)
                                if (t.CustomerID == t2.CustomerID)
                                {
                                    int j = 1;
                                    xlSheet1.Cells[i, j] = checkForEmpty(t.OrderID.ToString());
                                    j++;
                                    xlSheet1.Cells[i, j] = checkForEmpty(t2.ContactName.ToString());
                                    j++;
                                    // extra space to give uniform output in excel file
                                    xlSheet1.Cells[i, j] = " " + checkForEmpty(t.OrderDate.ToShortDateString());
                                    if (!unshippedOrders)
                                    {
                                        j++;
                                        xlSheet1.Cells[i, j] = " " + checkForEmpty(t.ShippedDate.ToShortDateString());
                                    }
                                    i++;
                                }
                    }
                    catch (StrongTypingException ex)
                    {
                        MessageBox.Show("Strong typing ex:\n" + ex);
                    }
                    
                }
                xlSheet1.Columns.AutoFit();
                if (unshippedOrders)
                    xlRange.Sort(xlRange.Columns[3], Excel.XlSortOrder.xlDescending);
                else
                    xlRange.Sort(xlRange.Columns[4], Excel.XlSortOrder.xlDescending);   
                
                //save
                xlBook.Close(true, Application.StartupPath + "\\" + saveFileAs, Type.Missing);
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString(), "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                //Release Resources and check file was created
                xlRange = null;
                xlSheet1 = null;
                xlBook = null;
                xlApp.Quit();
                xlApp = null;
                if(File.Exists(saveFileAs))
                    MessageBox.Show("Report exported to:\n" + Directory.GetCurrentDirectory() + "\\" + saveFileAs);
            }
        }

        /**
         * Menu needs it to work
         */
        private void exportAsToolStripMenuItem_Click(object sender, EventArgs e){}

        /**
         * Checks for changes to column values and checks if they are valid
         */
        private void Column_Changing(object sender, DataColumnChangeEventArgs args)
        {
            try
            {
                if (args.Column.ColumnName == "OrderID")
                {
                    args.Row.SetColumnError("OrderID", Convert.ToInt32(args.ProposedValue) <= 0 ? "Invalid value cannot be negative or empty" : "");
                }
                else if (args.Column.ColumnName == "TrackingID")
                {
                    args.Row.SetColumnError("TrackingID", Convert.ToInt32(args.ProposedValue) < 0 ? "Invalid value cannot be negative" : "");
                }
                else if (args.Column.ColumnName == "CustomerID")
                {
                    args.Row.SetColumnError("CustomerID", args.ProposedValue.ToString().Length < 5 ? "Customer ID cannot be longer than 5 letters" : "");
                }
                else if (args.Column.ColumnName == "ShippedDate")
                {
                    try
                    {
                        DateTime tmpDate;
                        if ((args.ProposedValue.ToString().Length > 1) && (DateTime.TryParse(args.ProposedValue.ToString(), out tmpDate) == false))
                            args.Row.SetColumnError("ShippedDate", "Invalid value");
                    }
                    catch (Exception)
                    {
                        args.ProposedValue = "";
                        Column_Changing(sender, args);
                    }
                }
            }
            catch (InvalidCastException)
            {
                if (args.Column.ColumnName == "TrackingID")
                    return;
                args.ProposedValue = 0;
                Column_Changing(sender, args);
            }
        }

        /**
         * Saves the users modified data values
         */
        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.customersBindingSource.EndEdit();
            this.shippersBindingSource.EndEdit();
            this.ordersBindingSource.EndEdit();

            if (a2DataSet.HasChanges())
            {
                A2DataSet dsChanged = (A2DataSet)a2DataSet.GetChanges();
                if (!dsChanged.HasErrors)
                {
                    A2DataSet.OrdersDataTable deletedOrders = (A2DataSet.OrdersDataTable)a2DataSet.Orders.GetChanges(DataRowState.Deleted);
                    A2DataSet.OrdersDataTable newOrders = (A2DataSet.OrdersDataTable)a2DataSet.Orders.GetChanges(DataRowState.Added);
                    A2DataSet.OrdersDataTable modifiedOrders = (A2DataSet.OrdersDataTable)a2DataSet.Orders.GetChanges(DataRowState.Modified);
                    try
                    {
                        //Remove all deleted orders from the Orders table.
                        if (deletedOrders != null)
                            ordersTableAdapter.Update(deletedOrders);
                        //Update the Customers and Shippers table.
                        customersTableAdapter.Update(a2DataSet.Customers);
                        shippersTableAdapter.Update(a2DataSet.Shippers);
                        //Add new orders to the Orders table.
                        if (newOrders != null)
                            ordersTableAdapter.Update(newOrders);
                        //Update all modified Orders.
                        if (modifiedOrders != null)
                            ordersTableAdapter.Update(modifiedOrders);

                        this.a2DataSet.AcceptChanges();
                        MessageBox.Show("Saved.", "Changes Saved", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString(), "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    finally
                    {
                        if (deletedOrders != null)
                            deletedOrders.Dispose();
                        if (newOrders != null)
                            newOrders.Dispose();
                        if (modifiedOrders != null)
                            modifiedOrders.Dispose();
                    }
                }
            }
        }

        /**
         * Exports the shipped orders to an excel file
         */
        private void shippedOrdersToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string saveFileAs = "completed.xlsx";
            String[] rowHeadings = new String[] { "Order ID", "Contact Name", "Date Ordered", "Date Shipped" };
            exportToExcel(saveFileAs, rowHeadings, false);
        }
    }
}
