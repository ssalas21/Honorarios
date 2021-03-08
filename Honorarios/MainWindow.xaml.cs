using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;
using Honorarios.BLL;

namespace Honorarios
{
    /// <summary>
    /// Lógica de interacción para MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            System.IO.StreamWriter file = new System.IO.StreamWriter("D:\\Honorarios\\2020\\Todos.txt"); // Abrir el txt
            for (int i = 1; i <= 9; i++)
            {
                string path = "D:\\Honorarios\\2020\\" + i + "\\CERTIFICADOS";
                string[] Lista = Directory.GetFiles(path);
                for (int j = 0; j < Lista.Length; j++)
                {
                    Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

                    if (xlApp == null)
                    {
                        MessageBox.Show("Excel is not properly installed!!");
                    }

                    Excel.Application xlApp2;
                    Excel.Workbook xlWorkBook2;
                    Excel.Worksheet xlWorkSheet2;
                    Excel.Range range;                    
                    xlApp2 = new Excel.Application();
                    xlWorkBook2 = xlApp2.Workbooks.Open(Lista[j], 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                    xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(1);
                    range = xlWorkSheet2.UsedRange;
                    string rut = ((range.Cells[7, 2] as Excel.Range).Value2).ToString();
                    string nombre = ((range.Cells[5, 2] as Excel.Range).Value2).ToString();                    
                    string[] nombres = nombre.Split(' ');
                    string[] ruts = rut.Split('-');                                     
                    for (int count = 18; count <= 29; count++)
                    {                        
                        int monto = Convert.ToInt32(((range.Cells[count, 5] as Excel.Range).Value2));
                        if (monto > 0)
                        {                            
                            int mes = count - 17;
                            int retencion = Convert.ToInt32(((range.Cells[count, 6] as Excel.Range).Value2));
                            file.WriteLine(";"+ruts[0] + ";" + ruts[1] + ";" + nombres[0] + ";" + nombres[1] + ";" + mes + ";2020;" + monto + ";0;" + retencion + ";0;0;0;0;0;0;30");
                        }
                    }
                    xlWorkBook2.Close(false, null, null);
                    xlApp2.Quit();
                    Marshal.ReleaseComObject(xlWorkSheet2);
                    Marshal.ReleaseComObject(xlWorkBook2);
                    Marshal.ReleaseComObject(xlApp2);
                }
            }
            file.Close();
            MessageBox.Show("Listo, siguiente paso a pasar a la DB!!!");
        }

        private void BtnArchivo_Click(object sender, RoutedEventArgs e)
        {
            System.IO.StreamWriter file = new System.IO.StreamWriter("D:\\Honorarios\\2020\\Paracargar.csv"); // Abrir el txt
            List<string> ruts = (new HonorariosBLL()).GetListRut();
            int count = 1;
            foreach (string item in ruts)
            {
                string rut = item;
                string digito = "";
                int total = 0;
                string enero = "";
                string febrero = "";
                string marzo = "";
                string abril = "";
                string mayo = "";
                string junio = "";
                string julio = "";
                string agosto = "";
                string septiembre = "";
                string octubre = "";
                string noviembre = "";
                string diciembre = "";
                List<HonorariosDatos> listado = (new HonorariosBLL()).GetBoletas(rut);
                foreach (HonorariosDatos aux in listado)
                {
                    digito = aux.Digito.ToString();
                    if (aux.Mes == 1) enero = "X";
                    if (aux.Mes == 2) febrero = "X";
                    if (aux.Mes == 3) marzo = "X";
                    if (aux.Mes == 4) abril = "X";
                    if (aux.Mes == 5) mayo = "X";
                    if (aux.Mes == 6) junio = "X";
                    if (aux.Mes == 7) julio = "X";
                    if (aux.Mes == 8) agosto = "X";
                    if (aux.Mes == 9) septiembre = "X";
                    if (aux.Mes == 10) octubre = "X";
                    if (aux.Mes == 11) noviembre = "X";
                    if (aux.Mes == 12) diciembre = "X";
                    total = total + Convert.ToInt32(aux.Descuento);
                }
                file.WriteLine(rut +";"+ digito + ";" + total + ";;;" + enero + ";" + febrero + ";" + marzo + ";" + abril + ";" + mayo + ";" + junio + ";" + julio + ";" + agosto + ";" + septiembre + ";" + octubre + ";" + noviembre + ";" + diciembre + ";;;" + count);
                count++;
            }
            file.Close();
            MessageBox.Show("Terminado, pasar a TXT el archivo y listo!!");
        }
    }
}
