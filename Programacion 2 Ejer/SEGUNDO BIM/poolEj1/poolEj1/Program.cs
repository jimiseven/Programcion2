using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;

namespace poolEj1
{
    class Program
    {
        static void Main(string[] args)
        {
            // Número de tareas a ejecutar
            int numberOfTasks = 10;

            Console.WriteLine("Iniciando tareas utilizando Thread Pool...");

            for (int i = 0; i < numberOfTasks; i++)
            {
                // Cola de tareas en el Thread Pool
                ThreadPool.QueueUserWorkItem(new WaitCallback(ProcessTask), i);
            }

            // Esperar a que el usuario presione una tecla para finalizar el programa
            Console.ReadLine();
        }

        // Método que representa la tarea a ser ejecutada por cada hilo en el pool
        static void ProcessTask(object taskId)
        {
            int id = (int)taskId;
            Console.WriteLine($"Tarea {id} comenzando en hilo {Thread.CurrentThread.ManagedThreadId}...");

            // Simular procesamiento con un retardo
            Thread.Sleep(2000);

            Console.WriteLine($"Tarea {id} completada en hilo {Thread.CurrentThread.ManagedThreadId}.");
        }
    }
}
