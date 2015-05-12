using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TXTextControl;

namespace WindowsFormsApplication11
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void mergeToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            DataSet ds = new DataSet();
            ds.ReadXml("sample_db.xml", XmlReadMode.Auto);

            mailMerge1.Merge(ds.Tables["Sales_SalesOrderHeader"]);

            textControl1.Tables.GridLines = false;
        }

        private void mailMerge1_BlockRowMerged(object sender, TXTextControl.DocumentServer.MailMerge.BlockRowMergedEventArgs e)
        {
            byte[] data;

            // create a temporary ServerTextControl to access API
            using (ServerTextControl tx = new ServerTextControl())
            {
                tx.Create();
                tx.Load(e.MergedBlockRow, BinaryStreamType.InternalUnicodeFormat);
                
                if (tx.Tables.GetItem() == null)
                    return;

                Table table = tx.Tables.GetItem();

                // loop through row and remove empty rows
                foreach (TableRow row in table.Rows)
                {
                    row.Select();
                    if (tx.Selection.Length == table.Columns.Count)
                    {
                        tx.Selection.Length = 0;
                        table.Rows.Remove();
                    }
                }

                tx.Save(out data, BinaryStreamType.InternalUnicodeFormat);
            }
            
            // return the manipulated merged block
            e.MergedBlockRow = data;
        }
    }
}
