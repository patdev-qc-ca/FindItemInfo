using Microsoft.Data.SqlClient;
using Microsoft.Office.Interop.Excel;

namespace FindItemInfo_net9
{
    public partial class FenetrePrincipale : Form
    {
        public FenetrePrincipale()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string[] requete = textBox1.Text.Split(' ');
            string sql = "SELECT [IDProjet],[NumItem],[Desc_FR],[Desc_EN],Manufact ,NomFournisseur FROM [autogrb].[dbo].[Projet_Pieces] " +
                "join Fournisseur on Projet_Pieces.IDFRS=Fournisseur.IDFRS " +
                "where ";
            for (int x = 0; x < requete.Length; x++)
            {
                if (x == requete.Length - 1)
                {
                    sql += "Desc_FR like '%ble%'";

                }
                else
                {
                    sql += "Desc_FR like '%ble%' or ";
                }
            }
            sql += " order by Desc_FR asc";
            SqlConnection con = new SqlConnection("Data Source=PRODUCTION\\SQLEXPRESS;Initial Catalog=AutoGRB;Integrated Security=false;Persist Security Info=True;User ID=sa; pwd=$NK#Bpl9YCm!0EKeZLdzp$Qetrz9g9bdQK7LO8L!u4oyv4rO2AOEvceyu8XIo;");
            con.Open();
            SqlDataReader rst = new SqlCommand(sql, con).ExecuteReader();
            while (rst.Read())
            {
                ListViewItem itm = listView1.Items.Add(string.Empty);
                itm.SubItems.Insert(0, new ListViewItem.ListViewSubItem(null, rst[0].ToString()));
                itm.SubItems.Insert(1, new ListViewItem.ListViewSubItem(null, rst[1].ToString()));
                itm.SubItems.Insert(2, new ListViewItem.ListViewSubItem(null, rst[2].ToString()));
                itm.SubItems.Insert(3, new ListViewItem.ListViewSubItem(null, rst[3].ToString()));
                itm.SubItems.Insert(4, new ListViewItem.ListViewSubItem(null, rst[4].ToString()));
                itm.SubItems.Insert(5, new ListViewItem.ListViewSubItem(null, rst[5].ToString()));
            }
            rst.Close();
            rst = null;
            con.Close();
            if (listView1.Items.Count <1)
            {
                MessageBox.Show("Données insuffisantes pour formuler une conclusion");
            }

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            string[] colones = {  "IDProjet","NumItem","Desc_FR","Desc_EN","Manufact","NomFournisseur"};
            foreach (string p in colones) {
                listView1.Columns.Add(p);
            }

        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {
            ListViewItem item = listView1.FocusedItem;
            MessageBox.Show($"IDProjet:{item.SubItems[0].Text}\nNumItem:{item.SubItems[1].Text}\nDesc_FR:{item.SubItems[2].Text}\nDesc_EN:{item.SubItems[3].Text}" +
                $"\nManufact:{item.SubItems[4].Text}\nNomFournisseur:{item.SubItems[5].Text}\n", this.Text);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application oXLApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook oXLBook = oXLApp.Workbooks.Add(string.Empty);
            Microsoft.Office.Interop.Excel.Worksheet oXLSheet = (Worksheet)oXLBook.Worksheets[1];
            oXLApp.Visible = true;
            oXLSheet.Range["A1: F1"].Font.Bold = true;
            oXLSheet.Name = "Recherche de pieces inventaire";
            oXLSheet.Range["A:F"].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            oXLSheet.Range["A1: F1"].Value = new object[] { "IDProjet", "NumItem", "Desc_FR", "Desc_EN", "Manufact", "NomFournisseur" };
            for (int X = 0; X < listView1.Items.Count; X++)
            {
                if (listView1.Items[X].Checked) {
                    oXLSheet.Range["A:F"].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    oXLSheet.Range[$"A{X + 2}: F{X + 2}"].Value = new object[]
                    {
                    listView1.Items[X].Text,
                    listView1.Items[X].SubItems[1].Text,
                    listView1.Items[X].SubItems[2].Text,
                    listView1.Items[X].SubItems[3].Text,
                    listView1.Items[X].SubItems[4].Text,
                    listView1.Items[X].SubItems[5].Text,
                    };
                }
            }
            oXLBook.SaveAs(Filename: $"{Environment.CurrentDirectory}\\Inventaire_{DateTime.Now.ToShortTimeString()}.xlsx");
            oXLBook.Close();
            oXLApp.Application.Quit();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            for (int a=0; a<listView1.Items.Count; a++) { listView1.Items[a].Checked = true; }
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            for (int a = 0; a < listView1.Items.Count; a++) { listView1.Items[a].Checked = false; }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            textBox1.Text = Clipboard.GetText();
            button1_Click(sender, e);
        }
    }
}
