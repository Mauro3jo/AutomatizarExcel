namespace Automatizacion_excel
{
    internal static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            // Preparar el entorno de la aplicaci�n (logs, config, etc.)
            PrepareAppEnvironment();

            // Manejo global de excepciones para robustez y trazabilidad
            Application.ThreadException += (sender, args) =>
            {
                // Aqu� podr�as integrar un logger real, por ejemplo: log4net, Serilog, NLog, etc.
                File.AppendAllText("error.log", $"{DateTime.Now}: ThreadException: {args.Exception}\n");
                MessageBox.Show("Ocurri� un error inesperado en el hilo principal.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            };

            AppDomain.CurrentDomain.UnhandledException += (sender, args) =>
            {
                File.AppendAllText("error.log", $"{DateTime.Now}: UnhandledException: {args.ExceptionObject}\n");
                MessageBox.Show("Ocurri� un error fatal. La aplicaci�n se cerrar�.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            };

            // Si tuvieras configuraci�n externa, podr�as cargarla aqu�
            // AppSettings.Load();

            // Inicio de la configuraci�n visual y arranque principal
            ApplicationConfiguration.Initialize();

            try
            {
                Application.Run(new Home());
            }
            catch (Exception ex)
            {
                File.AppendAllText("error.log", $"{DateTime.Now}: Exception in Main: {ex}\n");
                MessageBox.Show("Se produjo un error inesperado y la aplicaci�n debe cerrarse.", "Error cr�tico", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Preparaci�n del entorno de la aplicaci�n (carpetas, logs, etc.).
        /// </summary>
        private static void PrepareAppEnvironment()
        {
            // Ejemplo: Crear carpeta de logs si no existe
            string logDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logs");
            if (!Directory.Exists(logDir))
                Directory.CreateDirectory(logDir);

            // Redirigir logs a la carpeta 'logs'
            var logFilePath = Path.Combine(logDir, "error.log");
            if (!File.Exists(logFilePath))
                File.Create(logFilePath).Dispose();

            // Podr�as agregar m�s chequeos aqu�: config, recursos, etc.
        }
    }
}
