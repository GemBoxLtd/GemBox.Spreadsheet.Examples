using System;
using System.Data;
using System.Web.UI;
using System.Web.UI.WebControls;
using GemBox.Spreadsheet;

namespace MediumTrustCs
{
    public partial class _Default : Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

            // By specifying a location that is under ASP.NET application's control, 
            // GemBox.Spreadsheet can use file system operations to retrieve font data even in Medium Trust environment.
            FontSettings.FontsBaseDirectory = Server.MapPath("Fonts/");

            if (!Page.IsPostBack)
            {
                DataTable people = new DataTable();

                people.Columns.Add("ID", typeof(int));
                people.Columns.Add("FirstName", typeof(string));
                people.Columns.Add("LastName", typeof(string));

                Session["people"] = people;

                this.LoadDataFromFile(Request.PhysicalApplicationPath + "InData.xlsx");

                this.SetDataBinding();
            }
        }

        protected void Export_Click(object sender, EventArgs e)
        {
            DataTable people = (DataTable)Session["people"];

            // Create excel file.
            ExcelFile ef = new ExcelFile();
            ef.Styles.Normal.Font.Name = "Calibri";
            ExcelWorksheet ws = ef.Worksheets.Add("DataSheet");
            ws.InsertDataTable(people, new InsertDataTableOptions(0, 0) { ColumnHeaders = true });

            // Stream file to browser
            ef.Save(this.Response, "Report." + this.RadioButtonList1.SelectedValue);
        }

        private void LoadDataFromFile(string fileName)
        {
            DataTable people = (DataTable)Session["people"];

            ExcelFile ef = ExcelFile.Load(fileName);

            ExcelWorksheet ws = ef.Worksheets[0];

            ws.ExtractToDataTable(people, new ExtractToDataTableOptions("A1", ws.Rows.Count));
        }

        private void SetDataBinding()
        {
            DataTable people = (DataTable)Session["people"];
            DataView peopleDataView = people.DefaultView;

            this.GridView1.DataSource = peopleDataView;
            peopleDataView.AllowDelete = true;
            this.GridView1.DataBind();
        }

        protected void GridView1_RowDeleting(object sender, GridViewDeleteEventArgs e)
        {
            DataTable people = (DataTable)Session["people"];

            people.Rows[e.RowIndex].Delete();
            this.SetDataBinding();
        }

        protected void GridView1_RowEditing(object sender, GridViewEditEventArgs e)
        {
            this.GridView1.EditIndex = e.NewEditIndex;
            this.SetDataBinding();
        }

        protected void GridView1_RowUpdating(object sender, GridViewUpdateEventArgs e)
        {
            int i;
            int rowIndex = e.RowIndex;
            DataTable people = (DataTable)Session["people"];

            for (i = 1; i <= people.Columns.Count; i++)
            {
                TextBox editTextBox = this.GridView1.Rows[rowIndex].Cells[i].Controls[0] as TextBox;

                if (editTextBox != null)
                    people.Rows[rowIndex][i - 1] = editTextBox.Text;
            }

            this.GridView1.EditIndex = -1;
            this.SetDataBinding();
        }

        protected void GridView1_RowCancelingEdit(object sender, GridViewCancelEditEventArgs e)
        {
            this.GridView1.EditIndex = -1;
            this.SetDataBinding();
        }
    }
}