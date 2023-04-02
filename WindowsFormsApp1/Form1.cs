using SpreadsheetLight;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

// Cristian Andres Beltran Garzon
// Marzo de 2023

// vale, pues aparentemente los calculos ya funcionan (talvez con un margen de error de $ 3.000 en los intereses por mora)
// creo que ya se puede implementar la creacion del excel
// crear una tabla para mostrar los valores individuales
// indicar la fecha en que se actualizó el UVT y el % interes mora
/* a) Crear clase datos (los que quiero almacenar en DB)
 * exportar datos a XML
 * 
 * b) exportar datos y valores a la plantilla excel (necesito una clase?) 
 *      al excel solo hay que exportar 3 datos de calculos (ingresos, int ext e int mora), mas los
 *      datos personales
 * abrir la plantilla excel
 * probablemente se guarde automaticamente, en ese caso el nombre podría ser tal que Nombre_Apellido_01-01-2077_INDUCOM
 * 
 * hay que guardar el % mora cuando se modifique
 */
namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        private string plantillaExcel = (Environment.GetFolderPath(Environment.SpecialFolder.Desktop)+ @"\Inducom 2023\Plantilla Industria y Comercio.xlsx"); //@"C:\Users\User\source\repos\WindowsFormsApp1\WindowsFormsApp1\PlantillasExcel\Plantilla Industria y Comercio.xlsx";
        private double totalInteresMora = 0;
        private double totalInteresExtemporaneidad = 0;


        public Form1()
        {
            InitializeComponent();            
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            fechaLabel.Text = Program.fechaPagarGlobal.ToString("d");
            añoComboBox.Text = (Program.fechaPagarGlobal.Year - 1).ToString();
            dataGridView1.Visible = false;
            claseTb.Visible = false;
            claseBt.Visible = false;
        }
        
        private void textBox6_TextChanged(object sender, EventArgs e)              // Indica el total parcial en base a los ingresos, para que el usuario pueda comprobar que supera los 2 UVT
        {
            double d = (totalIngresosTb.Text != "") ? (double.Parse(totalIngresosTb.Text) * 0.006) : 0;
            controlTb.Text = (Math.Round(d / 1000) * 1000).ToString();          
        }

        private void button1_Click(object sender, EventArgs e)                     
        {
            double totalParcial = Math.Round(double.Parse((totalIngresosTb.Text != "") ? totalIngresosTb.Text : "0") * 0.006 / 1000) * 1000;  // Total parcial impuesto
          //  double totalParcial = Math.Round(double.Parse(totalIngresosTb.Text ?? "5"));
            double otrosCobros = Math.Round((totalParcial * 0.15) / 1000) * 1000;               // Cobros avisos y tablero
            otrosCobros += Math.Round((totalParcial * 0.05) / 1000) * 1000;                     // Cobro bomberil
                                                                                                                             
            double interesExtemporaneidad = 0, interesMora = 0;
           
            if (MesesTrasncurridos() >= 15){        // Verifica si el mes actual es abril o mayor, tomando como referencia los meses transcurridos desde enero del año Gravable
                interesExtemporaneidad = Math.Round((totalParcial * CalcularPorcExtemporaneidad() / 100) / 1000) * 1000;  //  Sancion de Extemporaneidad

                interesMora = (Math.Round((totalParcial * CalcularPorcMora() / 100) / 1000) * 1000);//  Intereses por mora
                interesMora = (interesMora != 0) ? interesMora : 1000;                              //  Si el valor fue redondeado a 0 por el calculo anterior, se redondeará automaticamente a 1000
            }
            //textBox1.Text = totalParcial.ToString();
            //textBox2.Text = interesExtemporaneidad.ToString();
            //textBox3.Text = interesMora.ToString();
            //textBox4.Text = otrosCobros.ToString();
            //textBox8.Text = CalcularPorcMora().ToString();
            totalInteresExtemporaneidad = interesExtemporaneidad;
            totalInteresMora = interesMora;
            extemTb.Text = interesExtemporaneidad.ToString();
            moraTb.Text = interesMora.ToString();

            totalPagarTb.Text = (totalParcial + otrosCobros + interesExtemporaneidad + interesMora).ToString();         //  Total a pagar
        
        }

        public double CalcularPorcExtemporaneidad()
        {            
            int mesesVencidos = MesesTrasncurridos() - 14;   // Cuantos meses han pasado desde marzo del año gravable
            return Math.Min(mesesVencidos * 5, 60);          // El porcentaje aumenta un 5% por mes vencido, máximo 60%                      
        }

        public double CalcularPorcMora() {
            double porcentajeDiario = (double.Parse(PorcMoraComboBox.Text)/12/30);  // Se calcula el porcentaje a pagar diario en base al porcentaje anual ingresado         
            double porcentajeMora;
            nombreTb.Text = porcentajeDiario.ToString();
            if (Program.fechaPagarGlobal.Year - int.Parse(añoComboBox.Text)<=1)     // Si el año a pagar es el inmediatamente anterior
            {
                porcentajeMora = (Program.fechaPagarGlobal.Day + ((MesesTrasncurridos() - 15) * 30)) * porcentajeDiario;   // Dias totales de mora multiplicados por el porcentaje diario                               
            }
            else                                                                    // Si hay mas de un año de mora
            {
                
                porcentajeMora = (Program.fechaPagarGlobal.Day + ((MesesTrasncurridos() - 16) * 30)) * porcentajeDiario;   //??// Dias totales de mora multiplicados por el porcentaje diario                               
            }
            cedulaTb.Text = porcentajeMora.ToString();
            return porcentajeMora; }
  
        private void button2_Click(object sender, EventArgs e)   // Seleccionar otra fecha
        {
            monthCalendar1.Visible = true;
        }

        private void monthCalendar1_DateChanged(object sender, DateRangeEventArgs e)
        {
            Program.fechaPagarGlobal = monthCalendar1.SelectionStart;
            fechaLabel.Text = Program.fechaPagarGlobal.ToString("d");            
        }
       
        public int MesesTrasncurridos()                         // Meses entre enero del año Gravable y la fecha de pago
        {
            DateTime añoPagar= new DateTime(int.Parse(añoComboBox.Text), 01, 01);                              
            return (Math.Abs((Program.fechaPagarGlobal.Month - añoPagar.Month) + 12 * (Program.fechaPagarGlobal.Year - añoPagar.Year)));
        }

        private void button3_Click(object sender, EventArgs e)  // Generar documento Excel listo para imprimir y firmar
        {
         if (plantillaExcel != "")
            {               
                SLDocument slDoc = new SLDocument(plantillaExcel); //por establecer
                
                slDoc.SetCellValue("E9", añoComboBox.Text);
                slDoc.SetCellValue("AC11", Program.fechaPagarGlobal.ToString("d"));
                slDoc.SetCellValue("K12", nombreTb.Text);
                slDoc.SetCellValue("L13", long.Parse((cedulaTb.Text != "") ? cedulaTb.Text : "0"));
                slDoc.SetCellValue("J14", direccionTb.Text);
                slDoc.SetCellValue("D17", long.Parse((celularTb.Text != "") ? celularTb.Text : "0"));
                slDoc.SetCellValue("I17", correoTb.Text);
                slDoc.SetCellValue("J28", codigoActividadTb.Text);
                slDoc.SetCellValue("AA18", long.Parse((totalIngresosTb.Text != "") ? totalIngresosTb.Text : "0"));
                slDoc.SetCellValue("AA45", totalInteresExtemporaneidad);
                slDoc.SetCellValue("AA51", totalInteresMora);
                slDoc.SetCellValue("J45", totalInteresExtemporaneidad == 0 ? "" : "X");

                string rutaNombreGuardar= ($@"C:\Users\User\Desktop\Inducom 2023\Inducom_{nombreTb.Text}_{DateTime.Now.ToString("MMMM-dd-yyyy  h mm tt")}.xlsx");  //nombre_2022 fecha hora.xlsx");   // por automatizar? !!                           
                slDoc.SaveAs(rutaNombreGuardar);
                Process.Start(rutaNombreGuardar);                                                                                                                                                                         //slDoc.SaveAs($@"C:\Users\User\source\repos\WindowsFormsApp1\WindowsFormsApp1\PlantillasExcel\Inducom_{nombreTb.Text}_{DateTime.Now.ToString("MMMM-dd-yyyy  h mm tt")}.xlsx");  //nombre_2022 fecha hora.xlsx");   // por automatizar? !!                           
            }   
        }

        private void claseBt_Click(object sender, EventArgs e)
        {
            Liquidacion primerDato = new Liquidacion(106913454, "nombre de ejemplo");

            claseTb.Text = primerDato.ToString();
            claseTb.Text = Environment.CurrentDirectory;
//          Environment.GetEnvironmentVariable()
        }
    }
    // RRG!! ÒʌÓ
}
