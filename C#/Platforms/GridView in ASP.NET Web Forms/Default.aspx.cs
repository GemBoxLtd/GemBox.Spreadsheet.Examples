using GemBox.Spreadsheet;
using System;
using System.Data;
using System.IO;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace AspNetGridView
{
    public partial class _Default : Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

            if (!Page.IsPostBack)
            {
                var people = new DataTable();
                people.Columns.Add("ID", typeof(int));
                people.Columns.Add("FirstName", typeof(string));
                people.Columns.Add("LastName", typeof(string));

                Session["people"] = people;

                this.LoadDataFromFile(Request.PhysicalApplicationPath + "InputData.xlsx");
                this.SetDataBinding();
            }
        }

        /// <summary>
        /// Export GridView data to Excel file.
        /// </summary>
        protected void ExportData_Click(object sender, EventArgs e)
        {
            var people = (DataTable)Session["people"];

            // Create Excel file.
            var workbook = new ExcelFile();
            var worksheet = workbook.Worksheets.Add("DataSheet");

            // Export DataTable that's used as GridView data source into Excel sheet.
            worksheet.InsertDataTable(people, new InsertDataTableOptions("A1") { ColumnHeaders = true });

            // Stream Excel file to client's browser.
            workbook.Save(this.Response, "Report." + this.RadioButtonList1.SelectedValue);
        }

        /// <summary>
        /// Export GridView data and formatting to Excel file.
        /// </summary>
        protected void ExportDataAndFormatting_Click(object sender, EventArgs e)
        {
            var stringWriter = new StringWriter();
            var htmlWriter = new HtmlTextWriter(stringWriter);
            
            // Export GridView control as HTML content.
            this.GridView1.RenderControl(htmlWriter);

            var htmlOptions = LoadOptions.HtmlDefault;
            var htmlData = htmlOptions.Encoding.GetBytes(stringWriter.ToString());

            using (var htmlStream = new MemoryStream(htmlData))
            {
                // Load HTML into Excel file.
                var workbook = ExcelFile.Load(htmlStream, htmlOptions);

                // Rename Excel sheet.
                var worksheet = workbook.Worksheets[0];
                worksheet.Name = "StyledDataSheet";

                // Delete Excel column that has Delete and Edit buttons.
                worksheet.Columns.Remove(0);

                // Stream Excel file to client's browser.
                workbook.Save(this.Response, "Styled Report." + this.RadioButtonList1.SelectedValue);
            }
        }

        // Override verification to successfully call GridView1.RenderControl method.
        public override void VerifyRenderingInServerForm(Control control)
        { }

        private void LoadDataFromFile(string fileName)
        {
            var people = (DataTable)Session["people"];

            // Load Excel file.
            var workbook = ExcelFile.Load(fileName);
            var worksheet = workbook.Worksheets[0];

            // Import Excel data into DataTable that's used as GridView data source.
            worksheet.ExtractToDataTable(people, new ExtractToDataTableOptions("A1", worksheet.Rows.Count));
        }

        private void SetDataBinding()
        {
            var people = (DataTable)Session["people"];
            var peopleDataView = people.DefaultView;
            peopleDataView.AllowDelete = true;

            this.GridView1.DataSource = peopleDataView;
            this.GridView1.DataBind();
        }

        protected void GridView1_RowDeleting(object sender, GridViewDeleteEventArgs e)
        {
            var people = (DataTable)Session["people"];
            people.Rows[e.RowIndex].Delete();
            this.SetDataBinding();
        }

        protected void GridView1_RowUpdating(object sender, GridViewUpdateEventArgs e)
        {
            var people = (DataTable)Session["people"];

            for (int i = 1; i <= people.Columns.Count; i++)
            {
                var editTextBox = this.GridView1.Rows[e.RowIndex].Cells[i].Controls[0] as TextBox;
                if (editTextBox != null)
                    people.Rows[e.RowIndex][i - 1] = editTextBox.Text;
            }

            this.GridView1.EditIndex = -1;
            this.SetDataBinding();
        }

        protected void GridView1_RowEditing(object sender, GridViewEditEventArgs e)
        {
            this.GridView1.EditIndex = e.NewEditIndex;
            this.SetDataBinding();
        }

        protected void GridView1_RowCancelingEdit(object sender, GridViewCancelEditEventArgs e)
        {
            this.GridView1.EditIndex = -1;
            this.SetDataBinding();
        }
    }
}