using SpreadsheetLight;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelSQLEscritorio.Models;

namespace ExcelSQLEscritorio
{
    public partial class btnCopiarSQL : Form
    {
        string[] cabeceras;
        private string path = @"C:\\";
        public btnCopiarSQL()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            string NombreHoja = tbHoja.Text;
            try
            {
                SLDocument sl = new SLDocument(path, NombreHoja);
                var lst = new List<Models.productosViewModel>();
                int iRow = 2; //Contador para administrar las filas, empezamos en 2 porque en 1 esta el encabezado
                int iCab = 1;
                while (!string.IsNullOrEmpty(sl.GetCellValueAsString(1, iCab)))
                {
                    iCab++;
                }
                cabeceras = new string[iCab-1];
                //cabeceras = new string[Convert.ToInt32(sl.GetColumnWidth(1))];
                int cab = cabeceras.Length;
                for (int i = 0; i < cabeceras.Length; i++) //Obtenemos las cabeceras que deben ser las mismas que en SQl
                {
                    int x = cabeceras.Length;
                    cabeceras[i] = sl.GetCellValueAsString(1, i + 1);
                }

                while (!string.IsNullOrEmpty(sl.GetCellValueAsString(iRow, 1)))
                {
                    productosViewModel oProducto = new productosViewModel
                    {
                        col1 = sl.GetCellValueAsString(iRow, 1),
                        col2 = sl.GetCellValueAsString(iRow, 2),
                        col3 = sl.GetCellValueAsString(iRow, 3),
                        col4 = sl.GetCellValueAsString(iRow, 4),
                        col5 = sl.GetCellValueAsString(iRow, 5),
                        col6 = sl.GetCellValueAsString(iRow, 6),
                        col7 = sl.GetCellValueAsString(iRow, 7),
                        col8 = sl.GetCellValueAsString(iRow, 8),
                        col9 = sl.GetCellValueAsString(iRow, 9),
                        col10 = sl.GetCellValueAsString(iRow, 10),
                        col11 = sl.GetCellValueAsString(iRow, 11),
                        col12 = sl.GetCellValueAsString(iRow, 12),
                        col13 = sl.GetCellValueAsString(iRow, 13),
                        col14 = sl.GetCellValueAsString(iRow, 14),
                        col15 = sl.GetCellValueAsString(iRow, 15),
                        col16 = sl.GetCellValueAsString(iRow, 16),
                        col17 = sl.GetCellValueAsString(iRow, 17),
                        col18 = sl.GetCellValueAsString(iRow, 18),
                        col19 = sl.GetCellValueAsString(iRow, 19),
                        col20 = sl.GetCellValueAsString(iRow, 20)
                    };

                    lst.Add(oProducto);

                    iRow++;
                }
                lblRegistros.Text = lst.Count.ToString() + " registros";
                dgvDatos.DataSource = lst;
            }
            catch(Exception ex)
            {
                MessageBox.Show("Selecciona un archivo o asegurate de que el archivo seleccionado este cerrado y no se esté utilizando", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnGenerar_Click(object sender, EventArgs e)
        {
            string codigo = "";
            string tempWhere = "";
            if (rbUpdate.Checked)
            {
                for (int i = 0; i < dgvDatos.Rows.Count; i++)
                {
                    codigo = codigo + "UPDATE " + tbTablaSQL.Text + "\n SET ";
                    for (int x = 0; x < cabeceras.Length; x++)
                    {
                        if (x == 0)
                        {
                            tempWhere = " WHERE " + cabeceras[x] + " = '" + dgvDatos.Rows[i].Cells[x].Value.ToString() + "' ";
                        }
                        else if (x == cabeceras.Length-1)
                        {
                            codigo = codigo + cabeceras[x] + " = " + dgvDatos.Rows[i].Cells[x].Value.ToString() + " ";
                        }
                        else
                        {
                            codigo = codigo + cabeceras[x] + " = " + dgvDatos.Rows[i].Cells[x].Value.ToString() + ", ";
                        }

                    }
                    codigo = codigo + "\n";
                    codigo = codigo + tempWhere + ";\n";
                }
            }
            else if (rbSelect.Checked)
            {
                for (int i = 0; i < dgvDatos.Rows.Count; i++)
                {
                    if (i == 0)
                    {
                        codigo = codigo + "SELECT * FROM " + tbTablaSQL.Text + "\n ";
                        for (int x = 0; x < cabeceras.Length; x++)
                        {
                            if (x == 0)
                            {
                                tempWhere = " WHERE " + cabeceras[x] + " = '" + dgvDatos.Rows[i].Cells[x].Value.ToString() + "'\n";
                            }

                        }
                        codigo = codigo + tempWhere;
                    }
                    else
                    {
                        for (int x = 0; x < cabeceras.Length; x++)
                        {
                            if (x == 0)
                            {
                                tempWhere = " OR " + cabeceras[x] + " = '" + dgvDatos.Rows[i].Cells[x].Value.ToString() + "'\n";
                            }

                        }
                        codigo = codigo + tempWhere;
                    }
                }
            }
            if (rbInsert.Checked)
            {
                tempWhere = tempWhere + "VALUES ('";
                for (int i = 0; i < dgvDatos.Rows.Count; i++)
                {
                    codigo = codigo + "INSERT INTO " + tbTablaSQL.Text + "\n(";
                    for (int x = 0; x < cabeceras.Length; x++)
                    {
                        if (x == cabeceras.Length - 1)
                        {
                            codigo = codigo + cabeceras[x] + ")\n";
                            tempWhere = tempWhere + dgvDatos.Rows[i].Cells[x].Value.ToString() + "')";
                        }
                        else
                        {
                            codigo = codigo + cabeceras[x] + ", ";
                            tempWhere = tempWhere + dgvDatos.Rows[i].Cells[x].Value.ToString() + "', '";
                        }
                    }

                    codigo = codigo + "\n";
                    codigo = codigo + tempWhere + ";\n";
                }
            }
            tbCodigoSQL.Text = codigo;
            
        }

        private void btnArchivo_Click(object sender, EventArgs e)
        {
            ofdExcel.InitialDirectory = "c:\\";
            ofdExcel.Filter = "Archivos Excel|*.xls;*.xlsx;*.xlsm";
            ofdExcel.FilterIndex = 2;
            ofdExcel.RestoreDirectory = true;
            if (ofdExcel.ShowDialog() == DialogResult.OK)
            {
                //Get the path of specified file
                path = ofdExcel.FileName;
            }
            lblRuta.Text = path;
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            if(tbCodigoSQL.Text.Length < 1)
            {
                MessageBox.Show("No hay contenido para copiar", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {
                Clipboard.SetText(tbCodigoSQL.Text);
                MessageBox.Show("Copiado al portapapeles", "Listo", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
        }
    }
}