using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.Diagnostics; // Agregar el espacio de nombres para Stopwatch

namespace poolEj2
{
    class Program
    {
        // Lista de materias
        static string[] subjects = {
            "Matemáticas", "Física", "Química", "Biología",
            "Historia", "Geografía", "Lengua", "Inglés",
            "Arte", "Informática","Matemáticas2", "Física2", "Química2", "Biología2",
            "Historia2", "Geografía2", "Lengua2", "Inglés2",
            "Arte2", "Informática2"
        };

        // SemaphoreSlim para limitar el número de hilos concurrentes
        //static SemaphoreSlim semaphore;

        // CountdownEvent para esperar a que todas las tareas se completen
        static CountdownEvent countdown;

        static void Main(string[] args)
        {
            int processorCount = Environment.ProcessorCount;
            Console.WriteLine("Número de hilos disponibles en el CPU: " + processorCount);
            Console.WriteLine("Iniciando estudio de materias utilizando Thread Pool...");

            // Inicializar el semáforo con el número de procesadores lógicos
            //semaphore = new SemaphoreSlim(processorCount);

            // Inicializar el CountdownEvent con el número de tareas
            countdown = new CountdownEvent(subjects.Length);

            // Crear un Stopwatch para medir el tiempo de ejecución
            //Stopwatch stopwatch = new Stopwatch();
            //stopwatch.Start();

            // Colocar cada materia en la cola de tareas del Thread Pool
            for (int i = 0; i < subjects.Length; i++)
            {
                ThreadPool.QueueUserWorkItem(new WaitCallback(StudySubject), subjects[i]);
            }

            // Esperar a que todas las tareas se completen
            countdown.Wait();

            // Detener el Stopwatch y mostrar el tiempo de ejecución
            //stopwatch.Stop();
            //Console.WriteLine("Tiempo total de ejecución: " + stopwatch.ElapsedMilliseconds + " ms");

            // Esperar a que el usuario presione una tecla para finalizar el programa
            Console.ReadLine();
        }

        // Método que representa la tarea de estudiar una materia
        static void StudySubject(object subject)
        {
            //semaphore.Wait(); // Esperar a que haya un recurso disponible en el semáforo
            try
                
            {
                string subjectName = (string)subject;
                Console.WriteLine($"+ Estudiando {subjectName} en hilo {Thread.CurrentThread.ManagedThreadId}...");

                // Simular tiempo de estudio con un retardo
                Thread.Sleep(new Random().Next(1000, 5000)); // Simular entre 1 y 5 segundos de estudio

                Console.WriteLine($"- Completado el estudio de {subjectName} en hilo {Thread.CurrentThread.ManagedThreadId}.");
            }
            finally
            {
                //semaphore.Release(); // Liberar el recurso del semáforo
                //countdown.Signal(); // Indicar que una tarea ha completado
            }
        }
    }
}
