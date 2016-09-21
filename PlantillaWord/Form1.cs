using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MySql.Data.MySqlClient;
using System.IO;
using AppAccounting;
namespace PlantillaWord
{
    public partial class Form1 : Form
    {
        private MySqlConnectionStringBuilder builder = new MySqlConnectionStringBuilder();
        public Form1()
        {
            InitializeComponent();
            builder.Server = "192.168.1.35";
            builder.UserID = "root";
            builder.Password = "admin";
            builder.Database = "empafrut";
        }

        private void btnimprimir_Click(object sender, EventArgs e)
        {
            string directorio = Application.StartupPath+"\\docs\\";
            string plantilla = "Plantilla.doc";
            string temporal = "PlantillaTemp.doc";
            if(validar()){
                try
                {
                    DateTime inicio;
                    DateTime fin;
                    if (DateTime.TryParse( txtinicio.Text, out inicio))
                    {
                        if (DateTime.TryParse(txtfin.Text, out fin))
                        {
                            WordDocument wd = new WordDocument(directorio, plantilla, temporal);
                            wd.FindAndReplace("<empleado>", txtempleado.Text);
                            wd.FindAndReplace("<dni>", txtdni.Text);
                            wd.FindAndReplace("<domicilio>", txtdir.Text);
                            wd.FindAndReplace("<cargo>", txtcargo.Text);
                            wd.FindAndReplace("<remuneracion>", txtremuneracion.Text);

                            int dias = 8;
                            int meses = 2;
                            wd.FindAndReplace("<meses>", meses);
                            wd.FindAndReplace("<dias>", dias);
                            wd.FindAndReplace("<inicio>", txtinicio.Text);
                            wd.FindAndReplace("<fin>", txtfin.Text);
                            DateTime fecha = DateTime.Now;
                            wd.FindAndReplace("<diaAct>", fecha.Day);
                            wd.FindAndReplace("<mesAct>", fecha.Month);
                            wd.FindAndReplace("<anioAct>", fecha.Year);
                            wd.SaveDocument();
                            wd.CloseDocument();
                            //lee el archivo
                            FileStream fs = File.Open(directorio + temporal, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                            Byte[] data = new byte[fs.Length];
                            fs.Read(data, 0, Convert.ToInt32(fs.Length));
                            fs.Close();
                            //conecta a la bd
                            MySqlConnection conn = new MySqlConnection(builder.ToString());
                            conn.Open();
                            using (var cmd = new MySqlCommand("INSERT INTO docs values( @nombre, @doc )", conn))
                            {
                                cmd.Parameters.Add("@nombre", MySqlDbType.VarChar).Value = txtdni.Text;
                                cmd.Parameters.Add("@doc", MySqlDbType.Blob).Value = data;
                                cmd.ExecuteNonQuery();
                            }
                            conn.Close(); //cierra conexion
                            MessageBox.Show("Listo");
                        }
                        else{MessageBox.Show("Formatos de Fecha de Fin Incorrecto");}
                    }else{MessageBox.Show("Formatos de Fecha de Inicio Incorrecto");} 
                }catch(Exception ex){
                    MessageBox.Show("Error: "+ex);
                }
            }
        }
        private bool validar()
        {
            if (txtdni.Text.Equals("") || txtempleado.Text.Equals("") || txtdir.Text.Equals("")
                || txtcargo.Text.Equals("") || txtcargo.Text.Equals("") || txtremuneracion.Text.Equals("")
                || txtinicio.Text.Equals("") || txtfin.Text.Equals(""))
            {
                MessageBox.Show("Campos Vacios");
                return false;
            }
            return true;
        }
        public double CalcularDias(DateTime primerFecha, DateTime segundaFecha)
        {
            TimeSpan diferencia;
            diferencia = primerFecha - segundaFecha;
            return Math.Abs(diferencia.Days);
        }
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                string directorio = Application.StartupPath + "\\WordDocs\\";
                string outputPath = @directorio + "\\ArchivoUnificado.doc";
                string template = directorio + "\\Template.dot";
                string temp1 = directorio + "\\Temp1.doc";
                if (File.Exists(outputPath))
                {
                    File.Delete(outputPath);
                }
                if (!File.Exists(temp1))
                {
                    File.Create(temp1).Close();
                }

                MySqlConnection conn = new MySqlConnection(builder.ToString());
                conn.Open();
                MySqlDataReader rdr;
                using (var cmd = new MySqlCommand("SELECT contrato FROM perlab Limit 1,7", conn))
                {
                    rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        if ( rdr[0] != DBNull.Value )
                        {
                            Byte[] data = (Byte[])rdr[0];
                            FileStream fs = File.Open(temp1, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite);
                            fs.Write(data, 0, data.Length);
                            fs.Close();
                            if (!File.Exists(outputPath))
                            {
                                File.Copy(temp1, outputPath);
                            }
                            else
                            {
                                WordMerge.Merge(new string[] {temp1, outputPath }, outputPath, true, template);
                            }
                            Console.Write("ad");
                        }
                    }
                }
                conn.Close();
                File.Delete(temp1);
                MessageBox.Show("Listo");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex);
            }
        }
    }
}
