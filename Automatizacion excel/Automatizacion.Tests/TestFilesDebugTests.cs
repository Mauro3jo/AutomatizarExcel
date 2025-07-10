using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;

namespace Automatizacion.Tests.Excel
{
    [TestClass]
    public class TestFilesDebugTests
    {
        [TestMethod]
        public void Debug_ListarArchivosEnTestFiles()
        {
            // Carpeta donde el test va a buscar los archivos (relativa al directorio de salida)
            string carpetaTestFiles = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles");
            System.Diagnostics.Debug.WriteLine($"Directorio actual: {Directory.GetCurrentDirectory()}");
            System.Diagnostics.Debug.WriteLine($"Carpeta buscada: {carpetaTestFiles}");
            System.Diagnostics.Debug.WriteLine("Archivos en TestFiles:");

            if (Directory.Exists(carpetaTestFiles))
            {
                foreach (var f in Directory.GetFiles(carpetaTestFiles))
                {
                    System.Diagnostics.Debug.WriteLine(f);
                }
            }
            else
            {
                System.Diagnostics.Debug.WriteLine("No existe la carpeta TestFiles en el output.");
            }
        }
    }
}
