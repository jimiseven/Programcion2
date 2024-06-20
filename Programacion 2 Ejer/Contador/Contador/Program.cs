using System;
using System.Diagnostics;
using System.IO;
using System.Threading;

namespace ProcesamientoArchivosConcurrente
{
    class Program
    {
        static void Main(string[] args)
        {
            string[] archivos = { "archivo1.txt", "archivo2.txt", "archivo3.txt" }; // Ejemplo de archivos

            Thread[] threads = new Thread[archivos.Length];

            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();

            // Crear y empezar los hilos
            for (int i = 0; i < archivos.Length; i++)
            {
                string archivo = archivos[i];
                threads[i] = new Thread(() => ProcesarArchivo(archivo));
                threads[i].Start();
            }

            // Esperar a que todos los hilos terminen
            foreach (Thread thread in threads)
            {
                thread.Join();
            }

            stopwatch.Stop();
            Console.WriteLine("Tiempo total de ejecución concurrente: " + stopwatch.ElapsedMilliseconds + " ms");
        }

        static void ProcesarArchivo(string rutaArchivo)
        {
            // Simular la lectura, procesamiento y guardado de datos del archivo
            Console.WriteLine($"Procesando {rutaArchivo}...");
            string contenido = File.ReadAllText(rutaArchivo);
            // Simular procesamiento (p. ej., contar palabras)
            int conteoPalabras = contenido.Split(' ').Length;
            // Guardar resultado (simulado)
            File.WriteAllText($"procesado_{rutaArchivo}", $"Conteo de palabras: {conteoPalabras}");
            Console.WriteLine($"Completado {rutaArchivo}");
        }
    }
}


