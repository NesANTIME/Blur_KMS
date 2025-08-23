using System;
using System.Diagnostics;
using System.IO;

namespace PythonLauncher
{
    class Program
    {
        static void Main()
        {
            string basePath = AppDomain.CurrentDomain.BaseDirectory;
            basePath = basePath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);

            string pythonPath = Path.Combine(basePath, "resources", "core", "interprete", "python-3.13.7-amd64", "python.exe");
            string scriptPath = Path.Combine(basePath, "Blur.py");

            if (!File.Exists(pythonPath))
            {
                Console.WriteLine($"ERROR: No se encuentra python.exe en: {pythonPath}");
                WaitForKey();
                return;
            }

            if (!File.Exists(scriptPath))
            {
                Console.WriteLine($"ERROR: No se encuentra Blur.py en: {scriptPath}");
                WaitForKey();
                return;
            }

            Console.WriteLine($"Python path: {pythonPath}");
            Console.WriteLine($"Script path: {scriptPath}");

            ProcessStartInfo psi = new ProcessStartInfo
            {
                FileName = pythonPath,
                Arguments = $"\"{scriptPath}\"",
                WorkingDirectory = basePath,
                UseShellExecute = false,  
                CreateNoWindow = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };

            try
            {
                using (Process process = new Process())
                {
                    process.StartInfo = psi;
                    process.OutputDataReceived += (sender, e) => {
                        if (!string.IsNullOrEmpty(e.Data))
                            Console.WriteLine($"PYTHON OUTPUT: {e.Data}");
                    };
                    process.ErrorDataReceived += (sender, e) => {
                        if (!string.IsNullOrEmpty(e.Data))
                            Console.WriteLine($"PYTHON ERROR: {e.Data}");
                    };

                    if (process.Start())
                    {
                        process.BeginOutputReadLine();
                        process.BeginErrorReadLine();

                        Console.WriteLine($"Proceso terminado con código: {process.ExitCode}");
                    }
                    else
                    {
                        Console.WriteLine("No se pudo iniciar el proceso");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error al iniciar Python: " + ex.Message);
            }
        }

        static void WaitForKey()
        {
            Console.WriteLine("\nPresione cualquier tecla para salir...");
            Console.ReadKey();
        }
    }
}