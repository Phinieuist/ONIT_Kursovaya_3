using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Win32;
using ONIT_Kurs_3.Entities;
using Spire.Doc;
using Excel = Microsoft.Office.Interop.Excel;


namespace ONIT_Kurs_3
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        Database DB;

        public MainWindow()
        {
            InitializeComponent();

            DB = new Database();
        }

        private void BtnCategories_Click(object sender, RoutedEventArgs e)
        {
            DataGrid.ItemsSource = DB.categories;
            DataGrid.Items.Refresh();
            Label1.Content = "Таблица Categories | Количество записей: " + DB.categories.Count;
        }

        private void BtnCustomerDemographics_Click(object sender, RoutedEventArgs e)
        {
            DataGrid.ItemsSource = DB.customerDemographics;
            DataGrid.Items.Refresh();
            Label1.Content = "Таблица CustomerDemographics | Количество записей: " + DB.customerDemographics.Count;
        }

        private void BtnCustomers_Click(object sender, RoutedEventArgs e)
        {
            DataGrid.ItemsSource = DB.customers;
            DataGrid.Items.Refresh();
            Label1.Content = "Таблица Customers | Количество записей: " + DB.customers.Count;
        }

        private void BtnEmployees_Click(object sender, RoutedEventArgs e)
        {
            DataGrid.ItemsSource = DB.employees;
            DataGrid.Items.Refresh();
            Label1.Content = "Таблица Employees | Количество записей: " + DB.employees.Count;
        }

        private void BtnEmployeeTerritories_Click(object sender, RoutedEventArgs e)
        {
            DataGrid.ItemsSource = DB.employeeTerritories;
            DataGrid.Items.Refresh();
            Label1.Content = "Таблица EmployeeTerritories | Количество записей: " + DB.employeeTerritories.Count;
        }

        private void BtnOrderDetails_Click(object sender, RoutedEventArgs e)
        {
            DataGrid.ItemsSource = DB.orderDetails;
            DataGrid.Items.Refresh();
            Label1.Content = "Таблица OrderDetails | Количество записей: " + DB.orderDetails.Count;
        }

        private void BtnOrders_Click(object sender, RoutedEventArgs e)
        {
            DataGrid.ItemsSource = DB.orders;
            DataGrid.Items.Refresh();
            Label1.Content = "Таблица Orders | Количество записей: " + DB.orders.Count;
        }

        private void BtnProducts_Click(object sender, RoutedEventArgs e)
        {
            DataGrid.ItemsSource = DB.products;
            DataGrid.Items.Refresh();
            Label1.Content = "Таблица Products | Количество записей: " + DB.products.Count;
        }

        private void BtnRegion_Click(object sender, RoutedEventArgs e)
        {
            DataGrid.ItemsSource = DB.regions;
            DataGrid.Items.Refresh();
            Label1.Content = "Таблица Region | Количество записей: " + DB.regions.Count;
        }

        private void BtnShippers_Click(object sender, RoutedEventArgs e)
        {
            DataGrid.ItemsSource = DB.shippers;
            DataGrid.Items.Refresh();
            Label1.Content = "Таблица Shippers | Количество записей: " + DB.shippers.Count;
        }

        private void BtnSuppliers_Click(object sender, RoutedEventArgs e)
        {
            DataGrid.ItemsSource = DB.suppliers;
            DataGrid.Items.Refresh();
            Label1.Content = "Таблица Suppliers | Количество записей: " + DB.suppliers.Count;
        }

        private void BtnTerritories_Click(object sender, RoutedEventArgs e)
        {
            DataGrid.ItemsSource = DB.territories;
            DataGrid.Items.Refresh();
            Label1.Content = "Таблица Territories | Количество записей: " + DB.territories.Count;
        }

        public DataTable ConvertToDataTable<T>(IList<T> data)
        {
            PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(typeof(T));
            DataTable table = new DataTable();
            foreach (PropertyDescriptor prop in properties)
                table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
            foreach (T item in data)
            {
                DataRow row = table.NewRow();
                foreach (PropertyDescriptor prop in properties)
                    row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;
                table.Rows.Add(row);
            }
            return table;
        }

        public static void ExportToExcel(DataTable DataTable, string ExcelFilePath = null)
        {
            Excel.Application Excel = null;

            try
            {
                int ColumnsCount;

                if (DataTable == null || (ColumnsCount = DataTable.Columns.Count) == 0)
                    throw new Exception("Экспорт в Excel: Null или пустая входная таблица!\n");

                // load excel, and create a new workbook
                Excel = new Microsoft.Office.Interop.Excel.Application();
                Excel.Workbooks.Add();

                // single worksheet
                Microsoft.Office.Interop.Excel._Worksheet Worksheet = Excel.ActiveSheet;

                object[] Header = new object[ColumnsCount];

                // column headings               
                for (int i = 0; i < ColumnsCount; i++)
                    Header[i] = DataTable.Columns[i].ColumnName;

                Microsoft.Office.Interop.Excel.Range HeaderRange = Worksheet.get_Range((Microsoft.Office.Interop.Excel.Range)(Worksheet.Cells[1, 1]), (Microsoft.Office.Interop.Excel.Range)(Worksheet.Cells[1, ColumnsCount]));
                HeaderRange.Value = Header;
                HeaderRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                HeaderRange.Font.Bold = true;
                HeaderRange.ColumnWidth = HeaderRange.ColumnWidth * 2;

                // DataCells
                int RowsCount = DataTable.Rows.Count;
                object[,] Cells = new object[RowsCount, ColumnsCount];

                for (int j = 0; j < RowsCount; j++)
                    for (int i = 0; i < ColumnsCount; i++)
                        Cells[j, i] = DataTable.Rows[j][i];

                Worksheet.get_Range((Microsoft.Office.Interop.Excel.Range)(Worksheet.Cells[2, 1]), (Microsoft.Office.Interop.Excel.Range)(Worksheet.Cells[RowsCount + 1, ColumnsCount])).Value = Cells;

                // check fielpath
                if (ExcelFilePath != null && ExcelFilePath != "")
                {
                    try
                    {
                        Worksheet.SaveAs(ExcelFilePath);

                        System.Windows.MessageBox.Show("Файл сохранён!", "Сохранение", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("Экспорт в Excel: файл Excel не может быть сохранён! Проверьте путь к файлу.\n"
                            + ex.Message);
                    }
                }
                else    // no filepath is given
                {
                    Excel.Visible = true;
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Экспорт в Excel: \n" + ex.Message);
            }
            finally
            {
                if (Excel != null)
                    Excel.Quit();
            }
        }

        private DataTable GetDataTable()
        {
            DataTable dataTable = null;

            switch (TableComboBox.SelectedIndex)
            {
                case 0: { dataTable = ConvertToDataTable(DB.categories); break; }
                case 1: { dataTable = ConvertToDataTable(DB.orderDetails); break; }
                case 2: { dataTable = ConvertToDataTable(DB.employeeTerritories); break; }
                case 3: { dataTable = ConvertToDataTable(DB.employees); break; }
                case 4: { dataTable = ConvertToDataTable(DB.customers); break; }
                case 5: { dataTable = ConvertToDataTable(DB.customerDemographics); break; }
                case 6: { dataTable = ConvertToDataTable(DB.territories); break; }
                case 7: { dataTable = ConvertToDataTable(DB.suppliers); break; }
                case 8: { dataTable = ConvertToDataTable(DB.shippers); break; }
                case 9: { dataTable = ConvertToDataTable(DB.regions); break; }
                case 10: { dataTable = ConvertToDataTable(DB.products); break; }
                case 11: { dataTable = ConvertToDataTable(DB.orders); break; }
            }

            return dataTable;
        }

        private void BtnExcel_Click(object sender, RoutedEventArgs e)
        {
            DataTable dataTable = GetDataTable();

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Файл excel (*.xlsx)|*.xlsx";
            if (saveFileDialog.ShowDialog() == true)
            {
                try
                {
                    ExportToExcel(dataTable, saveFileDialog.FileName);
                }
                catch(Exception ex) { System.Windows.MessageBox.Show(ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error); }
            }
        }

        private void BtnWord_Click(object sender, RoutedEventArgs e)
        {
            Document doc = null;
            try
            {
                doc = new Document();
                Spire.Doc.Table table = new Spire.Doc.Table(doc, true);
                
                DataTable dataTable = GetDataTable();

                table.AddRow();
                Spire.Doc.TableRow row = table.Rows[0];
                for (int jj = 0; jj < dataTable.Columns.Count; jj++)
                {
                    row.AddCell();
                    row.Cells[jj].AddParagraph().AppendText(dataTable.Columns[jj].ColumnName.ToString());
                }

                for (int i = 0; i < dataTable.Rows.Count; i++)
                {                  
                    table.AddRow();
                    row = table.Rows[i + 1];                       
                    for (int j = 0; j < dataTable.Columns.Count; j++)
                    {                    
                        row.Cells[j].AddParagraph().AppendText(dataTable.Rows[i][j].ToString());
                    }
                }
                doc.AddSection();
                doc.Sections[0].Tables.Add(table);
                //doc.AcceptChanges();

                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Файл word (*.docx)|*.docx";
                if (saveFileDialog.ShowDialog() == true)
                {
                    doc.SaveToFile(saveFileDialog.FileName, FileFormat.Docx2013);
                    MessageBox.Show("Файл сохранён!", "Сохранение", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                    
            }
            catch (Exception ex) { System.Windows.MessageBox.Show(ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error); }
            finally
            {
                if (doc != null)
                    doc.Dispose();
            }
        }
    }
}
