using CargaPlantillasCierreDeMes.Clases;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
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

namespace CargaPlantillasCierreDeMes
{
    /// <summary>
    /// Lógica de interacción para MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        private BackgroundWorker backgroundWorker;
        private CargarPlantillaCancelaProvisiones CargarPlantillaCancelaProvisiones;
        private CargaPlantillaProvisiones CargaPlantillaProvisiones;
        public MainWindow()
        {
            InitializeComponent();
            backgroundWorker = new BackgroundWorker
            {
                WorkerReportsProgress = true
            };
            backgroundWorker.DoWork += BackgroundWorker_DoWork;
            backgroundWorker.ProgressChanged += BackgroundWorker_ProgressChanged;
            backgroundWorker.RunWorkerCompleted += BackgroundWorker_RunWorkerCompleted;
        }

        private void CargarArchivo_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(txtRutaArchivo.Text))
            {
                MessageBox.Show("Debe seleccionar un archivo para cargarlo", "Validación", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (rbProvision.IsChecked.Value)
            {
                progressBar.Value = 0;
                progressBar.Visibility = Visibility.Visible;

                CargaPlantillaProvisiones = new CargaPlantillaProvisiones(txtRutaArchivo.Text, progress => backgroundWorker.ReportProgress(progress));

                //Se agregan parametros al backroundworker para mandar el valor de los controles al metodo doWork y no mande error de bloqueo de la aplicación
                backgroundWorker.RunWorkerAsync(new DoWorkArgs()
                {
                    tipoArchivo = DoWorkArgs.TipoArchivo.Provision
                });

            }
            else if (rbCancelaProvision.IsChecked.Value)
            {
                progressBar.Value = 0;
                progressBar.Visibility = Visibility.Visible;

                CargarPlantillaCancelaProvisiones = new CargarPlantillaCancelaProvisiones(txtRutaArchivo.Text, progress => backgroundWorker.ReportProgress(progress));

                //Se agregan parametros al backroundworker para mandar el valor de los controles al metodo doWork y no mande error de bloqueo de la aplicación
                backgroundWorker.RunWorkerAsync(new DoWorkArgs()
                {
                    tipoArchivo = DoWorkArgs.TipoArchivo.CancelaProvision
                });
            }
        }

        private void BuscarArchivo_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Archivos de imagen (*.xls;*.xlsx)|*.xls;*.xlsx";
            openFileDialog.Multiselect = false;

            if (openFileDialog.ShowDialog() == true)
            {
                txtRutaArchivo.Text = openFileDialog.FileName;
            }
        }

        private void BackgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                //El valor de los controles se toma de los argumento enviados desde el llamado del metodo backgroundWorker.RunWorkerAsync, para evitar bloqueo de la aplicación
                DoWorkArgs args = (DoWorkArgs)e.Argument;

                switch (args.tipoArchivo)
                {
                    case DoWorkArgs.TipoArchivo.Provision:
                        CargaPlantillaProvisiones.CargarPlantilla();
                        break;

                    case DoWorkArgs.TipoArchivo.CancelaProvision:
                        CargarPlantillaCancelaProvisiones.CargarPlantilla();
                        break;

                }

                // Reporta el progress                
                backgroundWorker.ReportProgress(100);
            }
            catch (Exception ex)
            {

                backgroundWorker.ReportProgress(0);

                //Dispatcher.Invoke evita que se bloquie el proceso mientras se ejecuta la accion, aqui lo utilizo para que no mande el error por llamar un constructor
                Dispatcher.Invoke(() =>
                {
                    MessageBox.Show(ex.Message, "Validación", MessageBoxButton.OK, MessageBoxImage.Warning);
                });
            }
        }

        private void BackgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            //Dispatcher.Invoke evita que se bloquie el proceso mientras se ejecuta la accion
            Dispatcher.Invoke(() =>
            {
                progressBar.Value = e.ProgressPercentage;
                switch (e.ProgressPercentage)
                {
                    case 0:
                        progressText.Text = "No se pudo concluir el proceso debido a que existen errores";
                        break;
                    case int n when (n > 0 && n <= 49):
                        progressText.Text = $"Leyendo archivo Excel %{n}";
                        break;
                    case int n when (n > 49 && n < 100):
                        progressText.Text = $"Guardando registros en BD %{n}";
                        break;
                    case 100:
                        progressText.Text = "El archivo se cargó correctamente";
                        break;

                }
            });
        }

        private void BackgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {

        }

        public class DoWorkArgs
        {

            public enum TipoArchivo
            {
                Provision,
                CancelaProvision
            }
            public TipoArchivo tipoArchivo { get; set; }
        }
    }
}
