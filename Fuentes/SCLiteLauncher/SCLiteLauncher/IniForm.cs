using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SCLite
{
    public partial class IniForm : Form
    {
        System.Windows.Forms.Timer t = new System.Windows.Forms.Timer();
        bool fadeIn = true;

        public IniForm()
        {
            InitializeComponent();
            ExtraFormSettings();
            // If we use solution2 we need to comment the following line.
            SetAndStartTimer();
        }

        private void SetAndStartTimer()
        {
            t.Interval = 100;
            t.Tick += new EventHandler(t_Tick);
            t.Start();
        }

        private void ExtraFormSettings()
        {
            this.FormBorderStyle = FormBorderStyle.None;
            this.Opacity = 0.5;
            this.BackgroundImage = SCLiteLaucher.Properties.Resources.LoaderSCLite;
        }


        void t_Tick(object sender, EventArgs e)
        {
            // Fade in by increasing the opacity of the splash to 1.0
            if (fadeIn)
            {
                if (this.Opacity < 1.0)
                {
                    this.Opacity += 0.1;
                }
                // After fadeIn complete, begin fadeOut
                else
                {
                    fadeIn = false;
                }
            }

            // After fadeIn and fadeOut complete, stop the timer and close this splash.
            if (!(fadeIn))
            {
                //Ocultamos el proceso actual
                this.Hide();
                //Proceso de Actualizacion
                System.Diagnostics.Process UpdateProcess = new System.Diagnostics.Process();
                // ubicacion donde esta el ejecutable
                UpdateProcess.StartInfo.WorkingDirectory = "C:\\SCLite";
                UpdateProcess.StartInfo.FileName = "SCLiteUpdate.exe"; // nombre del archivo a ejecutar con su extension
                UpdateProcess.Start(); // inicia el ejecutable
                UpdateProcess.WaitForExit(); // esta opción indica que el programa solo podra seguir cuando se cierre el ejecutable
                UpdateProcess.Close(); // cierra el ejecutable
                UpdateProcess.Dispose(); // libera memoria en la aplicacion

                //Proceso Normal
                System.Diagnostics.Process NormalProcess = new System.Diagnostics.Process();
                // ubicacion donde esta el ejecutable
                NormalProcess.StartInfo.WorkingDirectory = "C:\\SCLite";
                NormalProcess.StartInfo.FileName = "SCLite.exe"; // nombre del archivo a ejecutar con su extension
                NormalProcess.Start(); // inicia el ejecutable
                NormalProcess.WaitForExit(); // esta opción indica que el programa solo podra seguir cuando se cierre el ejecutable
                NormalProcess.Close(); // cierra el ejecutable
                NormalProcess.Dispose(); // libera memoria en la aplicacion
                t.Stop();
                this.Close();
            }
        }
    }
}
