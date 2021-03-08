using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Poplu
{
    public partial class Reminders : Form
    {
        public Reminders()
        {
            InitializeComponent();
        }

        private void Reminders_Load(object sender, EventArgs e)
        {

            var employess = ReadExcel();

            var todayDate = DateTime.Today;
            DatGridView_Load(new DataGridView(), employess, todayDate, new ControlConfiguration { LabelText = "Today Events", LabelXLocation = 0, LabelYLocation = 20, LabelBackgroundColor = Color.Black, LabelForColor = Color.Yellow, GridYLocation = 50 });

            var tomorrowDate = DateTime.Today.AddDays(1);
            DatGridView_Load(new DataGridView(), employess, tomorrowDate, new ControlConfiguration { LabelText = "Tomorrow Events", LabelXLocation = 0, LabelYLocation = 200, LabelBackgroundColor = Color.White, LabelForColor = Color.Green, GridYLocation = 230 });

            var yesterdayDate = DateTime.Today.AddDays(-1);
            DatGridView_Load(new DataGridView(), employess, yesterdayDate, new ControlConfiguration { LabelText = "Yesterday Events", LabelXLocation = 0, LabelYLocation = 400, LabelBackgroundColor = Color.White, LabelForColor = Color.Green, GridYLocation = 430 });
           
            //We have put this on end so user can see a page meesage when everything is ready
            MessageBox.Show("Hey, It's Poplu Time, Richa!! Lets see whats up today.", "Poplu Reminder");

        }

        public List<Employee> ReadExcel()
        {
            var employess = new List<Employee>();

            try
            {
                Excel.Application excelApp = new Excel.Application();
                if (excelApp != null)
                {
                    var baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
                    var filePath = Path.Combine(baseDirectory, "poplu-reminder.xlsx");

                    Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(filePath, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                    Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelWorkbook.Sheets[1];

                    Excel.Range excelRange = excelWorksheet.UsedRange;
                    int rowCount = excelRange.Rows.Count;
                    int colCount = excelRange.Columns.Count;

                    for (int i = 1; i <= rowCount; i++)
                    {
                        if (i != 1)
                        {
                            var name = (excelWorksheet.Cells[i, 1] as Excel.Range)?.Value.ToString();
                            var dobd = (excelWorksheet.Cells[i, 2] as Excel.Range)?.Value.ToString();
                            var dobm = (excelWorksheet.Cells[i, 3] as Excel.Range)?.Value.ToString();
                            var dojd = (excelWorksheet.Cells[i, 4] as Excel.Range)?.Value.ToString();
                            var dojm = (excelWorksheet.Cells[i, 5] as Excel.Range)?.Value.ToString();
                            var totalExp = (excelWorksheet.Cells[i, 6] as Excel.Range)?.Value.ToString();
                            var expYearAsOn = (excelWorksheet.Cells[i, 7] as Excel.Range)?.Value.ToString();

                            employess.Add(new Employee
                            {
                                Name = name,
                                DOBD = Convert.ToInt32(dobd),
                                DOBM = Convert.ToInt32(dobm),
                                DOJD = Convert.ToInt32(dojd),
                                DOJM = Convert.ToInt32(dojm),
                                TotalExp = Convert.ToInt32(totalExp),
                                ExpYearAsOn = Convert.ToInt32(expYearAsOn)
                            });
                        }
                    }

                    excelWorkbook.Close();
                    excelApp.Quit();
                    return employess;
                }
            }
            catch (Exception ex)
            {
                return null;
            }

            return employess;
        }

        private void DatGridView_Load(DataGridView myNewGrid, List<Employee> employees, DateTime eventDate ,ControlConfiguration controlConfiguration)
        {
            var label = new Label();
            label.Text = controlConfiguration.LabelText;
            label.ForeColor = Color.Yellow;
            label.Location = new System.Drawing.Point(controlConfiguration.LabelXLocation, controlConfiguration.LabelYLocation);
            label.BackColor = Color.Black;
            label.Width = 500;
            label.Font = new Font(label.Font.Name, 15, FontStyle.Bold | FontStyle.Underline); ;
            this.Controls.Add(label);

            myNewGrid = new System.Windows.Forms.DataGridView();
            ((System.ComponentModel.ISupportInitialize)(myNewGrid)).BeginInit();
            this.SuspendLayout();
            myNewGrid.Parent = this;  // You have to set the parent manually so that the grid is displayed on the form
            myNewGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            myNewGrid.Location = new System.Drawing.Point(10, controlConfiguration.GridYLocation);  // You will need to calculate this postion based on your other controls.  
            myNewGrid.Name = "myNewGrid";
            myNewGrid.Size = new System.Drawing.Size(1000, 100);  // You said you need the grid to be 12x12.  You can change the size here.
            myNewGrid.TabIndex = 0;
            myNewGrid.ColumnHeadersVisible = false; // You could turn this back on if you wanted, but this hides the headers that would say, "Cell1, Cell2...."
            myNewGrid.RowHeadersVisible = false;
            myNewGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            myNewGrid.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCells;
            myNewGrid.Font = new Font(label.Font.Name, 8, FontStyle.Bold); ;
            //myNewGrid.CellClick += MyNewGrid_CellClick;  // Set up an event handler for CellClick.  You handle this in the MyNewGrid_CellClick method, below
            ((System.ComponentModel.ISupportInitialize)(myNewGrid)).EndInit();
            this.ResumeLayout(false);
            myNewGrid.Visible = true;

            LoadGridData(myNewGrid, employees, eventDate);
        }

        private void LoadGridData(DataGridView myNewGrid, List<Employee> employees, DateTime eventDate)
        {
            var events = new List<EmployeeEvent>();
            
            var filteredEmployees = employees?.Where(e => (e.DOBD == eventDate.Day && e.DOBM == eventDate.Month) ||
                                             (e.DOJD == eventDate.Day && e.DOJM == eventDate.Month))
                                             .ToList();

            if(!employees?.Any()?? false)
            {
                events.Add(new EmployeeEvent { EmployeeName = "No Events Found!!"});
            }

            foreach (var employee in filteredEmployees)
            {
                var totalExperience = GetTotalYearsOfExperiece(employee.TotalExp, employee.ExpYearAsOn);

                if ((employee.DOBD == eventDate.Day && employee.DOBM == eventDate.Month) && (employee.DOJD == eventDate.Day && employee.DOJM == eventDate.Month))
                {
                    events.Add(new EmployeeEvent { EmployeeName = employee.Name, EventDate = $"{employee.DOBD}/{employee.DOBM}", EventName = $"Wowww!! Birthday and Anniversary both, it's {totalExperience} year completed" });
                } 
                else if (employee.DOBD == eventDate.Day && employee.DOBM == eventDate.Month)
                {
                    events.Add(new EmployeeEvent { EmployeeName = employee.Name, EventDate = $"{employee.DOBD}/{employee.DOBM}", EventName = "Happy Birthday!!" });
                }
                else
                {
                    events.Add(new EmployeeEvent { EmployeeName = employee.Name, EventDate = $"{employee.DOJD}/{employee.DOJM}", EventName = $"Anniversary Time!! it's {totalExperience} Year completed" });
                }
            }

            myNewGrid.DataSource = events;
        }

        private double GetTotalYearsOfExperiece(int exp, int expYearAsOn)
        {
            var curretYearDiffrence = DateTime.Now.Year - expYearAsOn;
            return curretYearDiffrence + exp;
        }

        private void MyNewGrid_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            //throw new NotImplementedException();
        }

    }
}
