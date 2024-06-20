using System;
using System.Diagnostics;
using System.Threading;

namespace NumerosPares
{
    class Program
    {
        static void Main(string[] args)
        {
            int totalPares = 1000;
            int numHilos = 12;
            int paresPorHilo = totalPares / numHilos;

            Thread[] threads = new Thread[numHilos];

            // Crear un Stopwatch para medir el tiempo de ejecución
            Stopwatch stopwatch = new Stopwatch();

            // Iniciar el Stopwatch
            stopwatch.Start();

            // Crear y empezar los hilos
            for (int i = 0; i < numHilos; i++)
            {
                int start = i * paresPorHilo;
                int end = (i + 1) * paresPorHilo;
                threads[i] = new Thread(() => MostrarNumerosPares(start, end));
                threads[i].Start();
            }

            // Esperar a que todos los hilos terminen
            foreach (Thread thread in threads)
            {
                thread.Join();
            }

            // Detener el Stopwatch
            stopwatch.Stop();

            // Mostrar el tiempo de ejecución
            Console.WriteLine("Tiempo total de ejecución: " + stopwatch.ElapsedMilliseconds + " ms");

            // Esperar a que el usuario presione una tecla para finalizar el programa
            Console.ReadLine();
        }

        static void MostrarNumerosPares(int start, int end)
        {
            for (int i = start; i < end; i++)
            {
                int numeroPar = i * 2;
                Console.WriteLine(numeroPar);
            }
        }
    }
}

